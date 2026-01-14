using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using AcCoreApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace CadFilesUpdater
{
    public class BlockInfo
    {
        public string BlockName { get; set; }
        public List<string> Attributes { get; set; }
        public string FilePath { get; set; }
        public string FileVersion { get; set; }

        public BlockInfo()
        {
            Attributes = new List<string>();
        }
    }

    public class BlockAnalyzer
    {
        // NOTE: AutoCAD operations must run on AutoCAD UI thread. We no longer marshal via ExecuteInApplicationContext
        // from a background thread because that caused cross-thread WPF exceptions / crashes in AutoCAD 2021.

        // NOTE:
        // In AutoCAD 2021, calling Editor.Command (QSAVE/ATTSYNC) from a modeless WPF UI event handler can throw
        // eInvalidInput. Therefore we avoid command-based save/sync and use API equivalents instead.

        public sealed class AttributeInstanceRow
        {
            public string FilePath { get; set; }
            public string LayoutName { get; set; } // "Model" or layout name
            public string BlockName { get; set; }
            public string BlockHandle { get; set; } // Handle string of BlockReference
            public Dictionary<string, string> Attributes { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        public readonly struct ChangeKey : IEquatable<ChangeKey>
        {
            public string FilePath { get; }
            public string BlockHandle { get; }
            public string BlockName { get; }
            public string AttributeTag { get; }

            public ChangeKey(string filePath, string blockHandle, string blockName, string attributeTag)
            {
                FilePath = filePath ?? "";
                BlockHandle = blockHandle ?? "";
                BlockName = blockName ?? "";
                AttributeTag = attributeTag ?? "";
            }

            public bool Equals(ChangeKey other) =>
                string.Equals(FilePath, other.FilePath, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(BlockHandle, other.BlockHandle, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(BlockName, other.BlockName, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(AttributeTag, other.AttributeTag, StringComparison.OrdinalIgnoreCase);

            public override bool Equals(object obj) => obj is ChangeKey other && Equals(other);

            public override int GetHashCode()
            {
                unchecked
                {
                    int hash = 17;
                    hash = hash * 23 + StringComparer.OrdinalIgnoreCase.GetHashCode(FilePath ?? "");
                    hash = hash * 23 + StringComparer.OrdinalIgnoreCase.GetHashCode(BlockHandle ?? "");
                    hash = hash * 23 + StringComparer.OrdinalIgnoreCase.GetHashCode(BlockName ?? "");
                    hash = hash * 23 + StringComparer.OrdinalIgnoreCase.GetHashCode(AttributeTag ?? "");
                    return hash;
                }
            }
        }

        public static List<AttributeInstanceRow> AnalyzeAttributeInstances(string filePath)
        {
            var results = new List<AttributeInstanceRow>();
            using (var db = new Database(false, true))
            {
                db.ReadDwgFile(filePath, FileOpenMode.OpenForReadAndAllShare, false, null);
                db.CloseInput(true);

                Database oldDb = HostApplicationServices.WorkingDatabase;
                HostApplicationServices.WorkingDatabase = db;
                try
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                        // Model space
                        var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                        ScanBlockTableRecordForAttributes(tr, ms, filePath, "Model", results);

                        // Paper space layouts
                        var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                        if (layoutDict != null)
                        {
                            foreach (DBDictionaryEntry entry in layoutDict)
                            {
                                var layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                                if (layout != null && layout.LayoutName != "Model")
                                {
                                    var ps = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                                    ScanBlockTableRecordForAttributes(tr, ps, filePath, layout.LayoutName, results);
                                }
                            }
                        }

                        tr.Commit();
                    }
                }
                finally
                {
                    HostApplicationServices.WorkingDatabase = oldDb;
                }
            }

            return results;
        }

        private static void ScanBlockTableRecordForAttributes(
            Transaction tr,
            BlockTableRecord btr,
            string filePath,
            string layoutName,
            List<AttributeInstanceRow> results)
        {
            foreach (ObjectId entId in btr)
            {
                var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                var br = ent as BlockReference;
                if (br == null) continue;

                try
                {
                    string blockName = null;
                    var dynId = br.DynamicBlockTableRecord;
                    if (dynId.IsValid)
                    {
                        var dynBtr = tr.GetObject(dynId, OpenMode.ForRead) as BlockTableRecord;
                        blockName = dynBtr?.Name;
                    }
                    if (string.IsNullOrWhiteSpace(blockName))
                    {
                        var staticBtr = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                        blockName = staticBtr?.Name ?? "";
                    }

                    var row = new AttributeInstanceRow
                    {
                        FilePath = filePath,
                        LayoutName = layoutName,
                        BlockName = blockName,
                        BlockHandle = br.Handle.ToString()
                    };

                    if (br.AttributeCollection != null)
                    {
                        foreach (ObjectId attId in br.AttributeCollection)
                        {
                            var attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                            if (attRef == null) continue;
                            if (string.IsNullOrWhiteSpace(attRef.Tag)) continue;
                            row.Attributes[attRef.Tag] = attRef.TextString ?? "";
                        }
                    }

                    if (row.Attributes.Count > 0)
                        results.Add(row);
                }
                catch
                {
                    // Best-effort scanning
                }
            }
        }

        public static UpdateResult SaveCachedChanges(
            IDictionary<ChangeKey, string> changes,
            List<string> filePaths,
            System.Action<int, int, string> progressCallback = null)
        {
            var result = new UpdateResult { TotalFiles = filePaths.Count };
            if (changes == null || changes.Count == 0)
                return result;

            // Group changes by file for efficiency.
            var byFile = new Dictionary<string, List<KeyValuePair<ChangeKey, string>>>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in changes)
            {
                if (!byFile.TryGetValue(kv.Key.FilePath, out var list))
                {
                    list = new List<KeyValuePair<ChangeKey, string>>();
                    byFile[kv.Key.FilePath] = list;
                }
                list.Add(kv);
            }

            foreach (var filePath in filePaths)
            {
                result.ProcessedFiles++;
                progressCallback?.Invoke(result.ProcessedFiles, result.TotalFiles, filePath);

                if (!byFile.TryGetValue(filePath, out var fileChanges) || fileChanges.Count == 0)
                {
                    result.SuccessfulFiles++;
                    continue;
                }

                try
                {
                    if (IsFileOpenInAutoCAD(filePath))
                    {
                        result.Errors.Add(new FileError(filePath,
                            "File is currently open in AutoCAD. Skipped saving to avoid conflicts. Please close the drawing and try again."));
                        result.FailedFiles++;
                        continue;
                    }

                    using (var db = new Database(false, true))
                    {
                        db.ReadDwgFile(filePath, FileOpenMode.OpenForReadAndAllShare, false, null);
                        db.CloseInput(true);

                        var originalVersion = db.OriginalFileVersion;

                        Database oldDb = HostApplicationServices.WorkingDatabase;
                        HostApplicationServices.WorkingDatabase = db;
                        try
                        {
                            using (var tr = db.TransactionManager.StartTransaction())
                            {
                                foreach (var kv in fileChanges)
                                {
                                    ApplySingleChange(tr, db, kv.Key, kv.Value);
                                }
                                tr.Commit();
                            }
                        }
                        finally
                        {
                            HostApplicationServices.WorkingDatabase = oldDb;
                        }

                        db.SaveAs(filePath, originalVersion);
                    }

                    result.SuccessfulFiles++;
                }
                catch (Autodesk.AutoCAD.Runtime.Exception acadEx)
                {
                    result.Errors.Add(new FileError(filePath, FormatAcadException(acadEx)));
                    result.FailedFiles++;
                }
                catch (Exception ex)
                {
                    result.Errors.Add(new FileError(filePath, ex.Message));
                    result.FailedFiles++;
                }
            }

            return result;
        }

        private static bool IsFileOpenInAutoCAD(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) return false;

            try
            {
                var dm = AcCoreApp.DocumentManager;
                if (dm == null) return false;

                foreach (Document doc in dm)
                {
                    try
                    {
                        var name = doc?.Name;
                        if (string.IsNullOrWhiteSpace(name)) continue;
                        if (string.Equals(name, filePath, StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                    catch
                    {
                        // ignore individual doc issues
                    }
                }
            }
            catch
            {
                // best-effort
            }

            return false;
        }

        private static string FormatAcadException(Autodesk.AutoCAD.Runtime.Exception acadEx)
        {
            if (acadEx == null) return "AutoCAD error";
            try
            {
                // Prefer explicit status to raw exception text, but avoid compile-time dependency
                // on enum member names (they differ between some AutoCAD .NET references).
                var statusName = acadEx.ErrorStatus.ToString();

                if (string.Equals(statusName, "eFileSharingViolation", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(statusName, "eFileShareViolation", StringComparison.OrdinalIgnoreCase))
                    return "File is in use (sharing violation). Close it in AutoCAD/another app and try again. (" + statusName + ")";

                if (string.Equals(statusName, "eFilerError", StringComparison.OrdinalIgnoreCase))
                    return "AutoCAD file I/O error. The file may be open/locked, read-only, or require recovery. (eFilerError)";

                if (string.Equals(statusName, "eAccessDenied", StringComparison.OrdinalIgnoreCase))
                    return "Access denied while reading/writing the file. Check permissions/read-only flag. (eAccessDenied)";

                return "AutoCAD error: " + statusName;
            }
            catch
            {
                return acadEx.Message ?? "AutoCAD error";
            }
        }

        private static void ApplySingleChange(Transaction tr, Database db, ChangeKey key, string newValue)
        {
            if (string.IsNullOrWhiteSpace(key.BlockHandle) || string.IsNullOrWhiteSpace(key.AttributeTag))
                return;

            try
            {
                var h = new Handle(Convert.ToInt64(key.BlockHandle, 16));
                var id = db.GetObjectId(false, h, 0);
                if (id == ObjectId.Null || !id.IsValid) return;

                var br = tr.GetObject(id, OpenMode.ForWrite) as BlockReference;
                if (br == null) return;

                // Determine effective block name and definition BTR
                BlockTableRecord defBtr = null;
                string effectiveName = null;

                var dynId = br.DynamicBlockTableRecord;
                if (dynId.IsValid)
                {
                    defBtr = tr.GetObject(dynId, OpenMode.ForRead) as BlockTableRecord;
                    effectiveName = defBtr?.Name;
                }
                if (defBtr == null)
                {
                    defBtr = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    effectiveName = defBtr?.Name;
                }

                if (!string.IsNullOrWhiteSpace(key.BlockName) &&
                    !string.IsNullOrWhiteSpace(effectiveName) &&
                    !effectiveName.Equals(key.BlockName, StringComparison.OrdinalIgnoreCase))
                {
                    return; // safety: don't update a different block than the one user edited
                }

                if (br.AttributeCollection == null) return;

                AttributeReference targetAtt = null;
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    var attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                    if (attRef == null) continue;
                    if (attRef.Tag != null && attRef.Tag.Equals(key.AttributeTag, StringComparison.OrdinalIgnoreCase))
                    {
                        targetAtt = attRef;
                        break;
                    }
                }
                if (targetAtt == null) return;

                // Update + ATTSYNC-style refresh from definition
                targetAtt.TextString = newValue ?? "";

                AttributeDefinition matchingAttDef = null;
                if (defBtr != null)
                {
                    foreach (ObjectId defId in defBtr)
                    {
                        var defEnt = tr.GetObject(defId, OpenMode.ForRead) as Entity;
                        if (defEnt is AttributeDefinition ad &&
                            ad.Tag.Equals(key.AttributeTag, StringComparison.OrdinalIgnoreCase))
                        {
                            matchingAttDef = ad;
                            break;
                        }
                    }
                }

                if (matchingAttDef != null)
                {
                    targetAtt.SetAttributeFromBlock(matchingAttDef, br.BlockTransform);
                    targetAtt.TextString = newValue ?? "";
                }

                try
                {
                    if (targetAtt.IsMTextAttribute)
                        targetAtt.UpdateMTextAttribute();
                }
                catch { }

                try { targetAtt.AdjustAlignment(db); } catch { }
                try { targetAtt.RecordGraphicsModified(true); } catch { }
                TryKickEntity(targetAtt);
            }
            catch
            {
                // Best-effort per attribute
            }
        }

        public static List<BlockInfo> AnalyzeFiles(List<string> filePaths)
        {
            System.Diagnostics.Debug.WriteLine($"[BlockAnalyzer] Rozpoczynam analizę {filePaths.Count} plików");
            var allBlocks = new List<BlockInfo>();

            foreach (var filePath in filePaths)
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine($"[BlockAnalyzer] Analizuję plik: {filePath}");
                    var blocks = AnalyzeFile(filePath);
                    System.Diagnostics.Debug.WriteLine($"[BlockAnalyzer] Znaleziono {blocks.Count} bloków w pliku {System.IO.Path.GetFileName(filePath)}");
                    
                    // Add all blocks from this file, preserving FilePath information
                    foreach (var block in blocks)
                    {
                        // Ensure FilePath is set
                        block.FilePath = filePath;
                        System.Diagnostics.Debug.WriteLine($"[BlockAnalyzer] Blok: {block.BlockName}, Atrybuty: {string.Join(", ", block.Attributes)}, Plik: {System.IO.Path.GetFileName(filePath)}");
                        allBlocks.Add(block);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[BlockAnalyzer] Błąd analizy pliku {filePath}: {ex.Message}");
                    System.Diagnostics.Debug.WriteLine($"[BlockAnalyzer] StackTrace: {ex.StackTrace}");
                }
            }

            System.Diagnostics.Debug.WriteLine($"[BlockAnalyzer] Łącznie znaleziono {allBlocks.Count} bloków (z wszystkich plików)");
            return allBlocks;
        }

        private static List<BlockInfo> AnalyzeFile(string filePath)
        {
            var blocks = new Dictionary<string, BlockInfo>();
            string fileVersion = null;

            using (var db = new Database(false, true))
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine($"[AnalyzeFile] Otwieram plik: {filePath}");
                    db.ReadDwgFile(filePath, FileOpenMode.OpenForReadAndAllShare, false, null);
                    fileVersion = db.OriginalFileVersion.ToString();
                    System.Diagnostics.Debug.WriteLine($"[AnalyzeFile] Wersja pliku: {fileVersion}");

                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        int blockTableCount = 0;
                        foreach (ObjectId id in bt) { blockTableCount++; }
                        System.Diagnostics.Debug.WriteLine($"[AnalyzeFile] Liczba bloków w BlockTable: {blockTableCount}");
                        
                        // Najpierw sprawdźmy wszystkie referencje bloków w model space
                        var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                        int msCount = 0;
                        foreach (ObjectId id in ms) { msCount++; }
                        System.Diagnostics.Debug.WriteLine($"[AnalyzeFile] Liczba obiektów w Model Space: {msCount}");
                        
                        var foundBlockRefs = new HashSet<string>();
                        
                        foreach (ObjectId entId in ms)
                        {
                            try
                            {
                                var ent = tr.GetObject(entId, OpenMode.ForRead);
                                if (ent is BlockReference br)
                                {
                                    var dynamicBtrId = br.DynamicBlockTableRecord;
                                    if (dynamicBtrId.IsValid)
                                    {
                                        var dynamicBtr = tr.GetObject(dynamicBtrId, OpenMode.ForRead) as BlockTableRecord;
                                        if (dynamicBtr != null && !dynamicBtr.IsAnonymous && !dynamicBtr.IsLayout)
                                        {
                                            var blockName = dynamicBtr.Name;
                                            foundBlockRefs.Add(blockName);
                                            System.Diagnostics.Debug.WriteLine($"[AnalyzeFile] Znaleziono referencję bloku dynamicznego: {blockName}");
                                            
                                            if (!blocks.ContainsKey(blockName))
                                            {
                                                blocks[blockName] = new BlockInfo
                                                {
                                                    BlockName = blockName,
                                                    FilePath = filePath,
                                                    FileVersion = fileVersion
                                                };
                                            }
                                            
                                            // Pobierz atrybuty z referencji bloku
                                            foreach (ObjectId attId in br.AttributeCollection)
                                            {
                                                var attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                                                if (attRef != null && !string.IsNullOrEmpty(attRef.Tag) &&
                                                    !blocks[blockName].Attributes.Contains(attRef.Tag))
                                                {
                                                    blocks[blockName].Attributes.Add(attRef.Tag);
                                                    System.Diagnostics.Debug.WriteLine($"[AnalyzeFile] Dodano atrybut: {attRef.Tag} do bloku {blockName}");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"[AnalyzeFile] Błąd podczas przetwarzania obiektu: {ex.Message}");
                            }
                        }
                        
                        // Teraz sprawdźmy definicje bloków dynamicznych
                        foreach (ObjectId btrId in bt)
                        {
                            try
                            {
                                var btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                                
                                if (btr != null && !btr.IsAnonymous && !btr.IsLayout)
                                {
                                    // Sprawdź czy to blok dynamiczny (może być zdefiniowany jako dynamiczny nawet jeśli nie ma referencji)
                                    bool isDynamic = btr.IsDynamicBlock;
                                    
                                    // Jeśli znaleźliśmy referencję tego bloku, to na pewno go dodaj
                                    if (foundBlockRefs.Contains(btr.Name) || isDynamic)
                                    {
                                        var blockName = btr.Name;
                                        
                                        if (!blocks.ContainsKey(blockName))
                                        {
                                            blocks[blockName] = new BlockInfo
                                            {
                                                BlockName = blockName,
                                                FilePath = filePath,
                                                FileVersion = fileVersion
                                            };
                                        }

                                        // Get attributes from block definition
                                        foreach (ObjectId objId in btr)
                                        {
                                            var obj = tr.GetObject(objId, OpenMode.ForRead);
                                            if (obj is AttributeDefinition attrDef)
                                            {
                                                if (!string.IsNullOrEmpty(attrDef.Tag) && 
                                                    !blocks[blockName].Attributes.Contains(attrDef.Tag))
                                                {
                                                    blocks[blockName].Attributes.Add(attrDef.Tag);
                                                    System.Diagnostics.Debug.WriteLine($"[AnalyzeFile] Dodano atrybut z definicji: {attrDef.Tag} do bloku {blockName}");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"[AnalyzeFile] Błąd podczas przetwarzania BlockTableRecord: {ex.Message}");
                            }
                        }

                        tr.Commit();
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[AnalyzeFile] Błąd odczytu pliku {filePath}: {ex.Message}");
                    System.Diagnostics.Debug.WriteLine($"[AnalyzeFile] StackTrace: {ex.StackTrace}");
                }
            }

            System.Diagnostics.Debug.WriteLine($"[AnalyzeFile] Zwracam {blocks.Count} bloków z pliku {System.IO.Path.GetFileName(filePath)}");
            return blocks.Values.ToList();
        }

        public static DwgVersion GetFileVersion(string filePath)
        {
            try
            {
                using (var db = new Database(false, true))
                {
                    db.ReadDwgFile(filePath, FileOpenMode.OpenForReadAndAllShare, false, null);
                    return db.OriginalFileVersion;
                }
            }
            catch
            {
                return DwgVersion.AC1027; // Default to AutoCAD 2013 format
            }
        }

        public static UpdateResult UpdateBlocksInFiles(
            List<string> filePaths,
            string blockName,
            string attributeName,
            string value,
            System.Action<int, int, string> progressCallback = null)
        {
            var result = new UpdateResult
            {
                TotalFiles = filePaths.Count
            };

            foreach (var filePath in filePaths)
            {
                try
                {
                    result.ProcessedFiles++;
                    System.Diagnostics.Debug.WriteLine($"[UpdateBlocksInFiles] Processing file {result.ProcessedFiles}/{result.TotalFiles}: {filePath}");
                    System.Diagnostics.Debug.WriteLine($"[UpdateBlocksInFiles] Current thread: {System.Threading.Thread.CurrentThread.ManagedThreadId}");
                    
                    // Invoke callback on UI thread if it's provided
                    if (progressCallback != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocksInFiles] Calling progress callback");
                        try
                        {
                            progressCallback(result.ProcessedFiles, result.TotalFiles, filePath);
                            System.Diagnostics.Debug.WriteLine($"[UpdateBlocksInFiles] Progress callback completed successfully");
                        }
                        catch (Exception callbackEx)
                        {
                            // Ignore callback errors (e.g., if window is closed)
                            System.Diagnostics.Debug.WriteLine($"[UpdateBlocksInFiles] Callback error: {callbackEx.Message}");
                            System.Diagnostics.Debug.WriteLine($"[UpdateBlocksInFiles] Callback error type: {callbackEx.GetType().FullName}");
                            System.Diagnostics.Debug.WriteLine($"[UpdateBlocksInFiles] Callback error StackTrace: {callbackEx.StackTrace}");
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocksInFiles] progressCallback is null, skipping");
                    }
                    
                    bool fileSucceeded = false;
                    bool updatedInFile = false;
                    string fileError = null;

                    try
                    {
                        // OFFLINE / SIDE DATABASE MODE:
                        // We do not open/switch AutoCAD documents and we cannot run commands (MOVE/ATTSYNC/REGEN).
                        // We directly edit the DWG database and save it back preserving the original DWG version.
                        using (var db = new Database(false, true))
                        {
                            db.ReadDwgFile(filePath, FileOpenMode.OpenForReadAndAllShare, false, null);
                            db.CloseInput(true);

                            var originalVersion = db.OriginalFileVersion;
                            System.Diagnostics.Debug.WriteLine($"[UpdateBlocksInFiles] Original DWG version on disk: {originalVersion} ({filePath})");

                            Database oldDb = HostApplicationServices.WorkingDatabase;
                            HostApplicationServices.WorkingDatabase = db;
                            try
                            {
                                using (var tr = db.TransactionManager.StartTransaction())
                                {
                                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                                    updatedInFile = false;

                                    // Model space
                                    var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                                    updatedInFile |= UpdateBlocksInBlockTableRecord(tr, db, ms, blockName, attributeName, value);

                                    // Paper space layouts
                                    var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                                    if (layoutDict != null)
                                    {
                                        foreach (DBDictionaryEntry entry in layoutDict)
                                        {
                                            var layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                                            if (layout != null && layout.LayoutName != "Model")
                                            {
                                                var ps = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);
                                                updatedInFile |= UpdateBlocksInBlockTableRecord(tr, db, ps, blockName, attributeName, value);
                                            }
                                        }
                                    }

                                    tr.Commit();
                                }
                            }
                            finally
                            {
                                HostApplicationServices.WorkingDatabase = oldDb;
                            }

                            if (updatedInFile)
                            {
                                // Save back preserving the file's original DWG version (do NOT use user's default).
                                db.SaveAs(filePath, originalVersion);
                            }

                            fileSucceeded = true;
                        }
                    }
                    catch (System.Exception exOuter)
                    {
                        fileError = exOuter.Message;
                    }

                    if (!fileSucceeded)
                    {
                        var msg = string.IsNullOrWhiteSpace(fileError) ? "Failed to update file" : fileError;
                        result.Errors.Add(new FileError(filePath, msg));
                        result.FailedFiles++;
                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocksInFiles] ERROR: {filePath}: {msg}");
                    }
                    else
                    {
                        // We count the file as successful even if no blocks matched (no changes),
                        // since the operation completed without errors.
                        result.SuccessfulFiles++;
                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocksInFiles] Completed {filePath}. updated={updatedInFile}");
                    }
                }
                catch (System.IO.IOException ioEx)
                {
                    string errorMsg = "File is in use by another application";
                    if (ioEx.Message.Contains("locked") || ioEx.Message.Contains("access"))
                    {
                        errorMsg = "File is locked or access denied";
                    }
                    result.Errors.Add(new FileError(filePath, errorMsg));
                    result.FailedFiles++;
                    System.Diagnostics.Debug.WriteLine($"Błąd aktualizacji pliku {filePath}: {errorMsg}");
                }
                catch (Autodesk.AutoCAD.Runtime.Exception acadEx)
                {
                    string errorMsg = acadEx.Message;
                    if (errorMsg.Contains("eFilerError"))
                    {
                        errorMsg = "File error - possibly corrupted or in use";
                    }
                    result.Errors.Add(new FileError(filePath, errorMsg));
                    result.FailedFiles++;
                    System.Diagnostics.Debug.WriteLine($"Błąd aktualizacji pliku {filePath}: {errorMsg}");
                }
                catch (Exception ex)
                {
                    result.Errors.Add(new FileError(filePath, ex.Message));
                    result.FailedFiles++;
                    System.Diagnostics.Debug.WriteLine($"Błąd aktualizacji pliku {filePath}: {ex.Message}");
                }
            }

            return result;
        }

        private static bool UpdateBlocksInBlockTableRecord(Transaction tr, Database db, BlockTableRecord btr, string blockName, string attributeName, string value)
        {
            bool updated = false;

            foreach (ObjectId entId in btr)
            {
                var ent = tr.GetObject(entId, OpenMode.ForWrite);
                if (ent is BlockReference br)
                {
                    try
                    {
                        var dynamicBtrId = br.DynamicBlockTableRecord;
                        if (dynamicBtrId.IsValid)
                        {
                            var dynamicBtr = tr.GetObject(dynamicBtrId, OpenMode.ForRead) as BlockTableRecord;
                            if (dynamicBtr != null &&
                                dynamicBtr.Name.Equals(blockName, StringComparison.OrdinalIgnoreCase))
                            {
                                foreach (ObjectId attId in br.AttributeCollection)
                                {
                                    var attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                                    if (attRef != null && attRef.Tag.Equals(attributeName, StringComparison.OrdinalIgnoreCase))
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks] ===== Updating attribute '{attributeName}' in block '{blockName}' =====");
                                        
                                        // Get original text
                                        var originalText = attRef.TextString ?? "";
                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks] Original text: '{originalText}' (length: {originalText.Length})");
                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks] New text: '{value}' (length: {value?.Length ?? 0})");
                                        
                                        // Preserve ALL original formatting properties
                                        var originalJustify = attRef.Justify;
                                        var originalHeight = attRef.Height;
                                        var originalWidthFactor = attRef.WidthFactor;
                                        var originalRotation = attRef.Rotation;
                                        var originalIsMirroredInX = attRef.IsMirroredInX;
                                        var originalIsMirroredInY = attRef.IsMirroredInY;
                                        var originalPosition = attRef.Position;
                                        
                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks] BEFORE CHANGE:");
                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   Justify={originalJustify}");
                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   Position={originalPosition}");
                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   Height={originalHeight}, WidthFactor={originalWidthFactor}, Rotation={originalRotation}");
                                        
                                        // Try to get AlignmentPoint - it may not be available for all justify types
                                        Point3d originalAlignmentPoint = Point3d.Origin;
                                        bool hasAlignmentPoint = false;
                                        try
                                        {
                                            originalAlignmentPoint = attRef.AlignmentPoint;
                                            hasAlignmentPoint = true;
                                            System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   AlignmentPoint={originalAlignmentPoint}");
                                        }
                                        catch
                                        {
                                            hasAlignmentPoint = false;
                                            System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   AlignmentPoint: N/A (left-aligned)");
                                        }

                                        // Apply new value and then re-sync attribute reference from the AttributeDefinition.
                                        // This mimics what happens when you "toggle justification" in Block Editor / ATTSYNC:
                                        // AutoCAD re-applies the AttributeDefinition geometry to the AttributeReference, which fixes the visual centering.
                                        Database oldDb = HostApplicationServices.WorkingDatabase;
                                        HostApplicationServices.WorkingDatabase = db;

                                        try
                                        {
                                            var isMTextAttribute = false;
                                            try
                                            {
                                                isMTextAttribute = attRef.IsMTextAttribute;
                                            }
                                            catch
                                            {
                                                // older objects / edge cases
                                            }

                                            System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   IsMTextAttribute={isMTextAttribute}");

                                            // 1) Change the value
                                            System.Diagnostics.Debug.WriteLine($"[UpdateBlocks] Changing TextString from '{originalText}' to '{value}'");
                                            attRef.TextString = value;

                                            // 2) Find matching AttributeDefinition (by tag) in the *dynamic* block definition
                                            AttributeDefinition matchingAttDef = null;
                                            foreach (ObjectId defId in dynamicBtr)
                                            {
                                                var defEnt = tr.GetObject(defId, OpenMode.ForRead) as Entity;
                                                if (defEnt is AttributeDefinition ad &&
                                                    ad.Tag.Equals(attributeName, StringComparison.OrdinalIgnoreCase))
                                                {
                                                    matchingAttDef = ad;
                                                    break;
                                                }
                                            }

                                            if (matchingAttDef != null)
                                            {
                                                System.Diagnostics.Debug.WriteLine($"[UpdateBlocks] Found AttributeDefinition for tag '{attributeName}'. Re-syncing AttributeReference from definition...");

                                                try
                                                {
                                                    // Helpful diagnostics: where does the definition place this attribute?
                                                    // (Definition coords are in block space; transform into current space)
                                                    var defPos = matchingAttDef.Position;
                                                    Point3d defAlign = Point3d.Origin;
                                                    bool defHasAlign = false;
                                                    try
                                                    {
                                                        defAlign = matchingAttDef.AlignmentPoint;
                                                        defHasAlign = true;
                                                    }
                                                    catch { }

                                                    var defPosWcs = defPos.TransformBy(br.BlockTransform);
                                                    System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   Def.Position(block)={defPos} -> world={defPosWcs}");
                                                    if (defHasAlign)
                                                    {
                                                        var defAlignWcs = defAlign.TransformBy(br.BlockTransform);
                                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   Def.AlignmentPoint(block)={defAlign} -> world={defAlignWcs}");
                                                    }
                                                }
                                                catch { }

                                                // 3) Sync geometry/justification/alignment from definition (ATTSYNC-style)
                                                attRef.SetAttributeFromBlock(matchingAttDef, br.BlockTransform);

                                                // 4) Set the value again (SetAttributeFromBlock may restore default text)
                                                attRef.TextString = value;

                                                // 4b) For MText attributes, TextString updates a backing MText object.
                                                // In "side database" mode this often needs an explicit refresh.
                                                if (isMTextAttribute)
                                                {
                                                    try
                                                    {
                                                        attRef.UpdateMTextAttribute();
                                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks] Updated MTextAttribute representation");
                                                    }
                                                    catch (Exception exMText)
                                                    {
                                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks] WARNING: UpdateMTextAttribute failed: {exMText.Message}");
                                                    }
                                                }

                                                // 5) Force alignment calculation
                                                // IMPORTANT: do NOT restore the previous AlignmentPoint here.
                                                // The old AlignmentPoint may already be "broken" (left-anchored-at-center symptom).
                                                // We want the definition geometry to fully take effect.
                                                attRef.AdjustAlignment(db);
                                                try { attRef.RecordGraphicsModified(true); } catch { }

                                                // 6) "Kick" (MOVE 0,0 -> 0,0 equivalent) in side-database:
                                                // We cannot run AutoCAD commands offline, so we simulate a tiny modify+undo on geometry.
                                                // This is best-effort and should be no-op in practice, but can trigger recalculation in some cases.
                                                TryKickEntity(attRef);
                                                System.Diagnostics.Debug.WriteLine($"[UpdateBlocks] Re-sync + AdjustAlignment done.");

                                                // Post-sync diagnostics
                                                try
                                                {
                                                    System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   Ref.Justify(after)={attRef.Justify}");
                                                    System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   Ref.Position(after)={attRef.Position}");
                                                    try { System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   Ref.AlignmentPoint(after)={attRef.AlignmentPoint}"); } catch { }
                                                }
                                                catch { }
                                            }
                                            else
                                            {
                                                System.Diagnostics.Debug.WriteLine($"[UpdateBlocks] WARNING: Could not find AttributeDefinition for tag '{attributeName}' in dynamic block definition. Skipping re-sync.");
                                            }
                                        }
                                        finally
                                        {
                                            HostApplicationServices.WorkingDatabase = oldDb;
                                        }

                                        // NOTE: We intentionally do NOT "restore" AlignmentPoint/Position based on previous values.
                                        // In many problematic cases the stored AlignmentPoint/Position is already wrong after a TextString change.
                                        // Our goal is to let AutoCAD's definition-driven geometry (SetAttributeFromBlock + AdjustAlignment) win.

                                        // Final values
                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks] FINAL VALUES:");
                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   Justify={attRef.Justify}");
                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   Position={attRef.Position}");
                                        if (hasAlignmentPoint)
                                        {
                                            try
                                            {
                                                System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   AlignmentPoint={attRef.AlignmentPoint}");
                                            }
                                            catch
                                            {
                                                System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   AlignmentPoint: ERROR reading");
                                            }
                                        }
                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks] ===== Update complete =====");
                                        
                                        updated = true;
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Skip invalid block references
                    }
                }
            }

            return updated;
        }

        private static void TryKickEntity(Entity ent)
        {
            try
            {
                // Tiny round-trip displacement to mark entity modified without changing its effective location.
                // Using a very small epsilon reduces the chance of visible movement / precision issues.
                const double eps = 1e-9;
                var v = new Vector3d(eps, 0, 0);
                ent.TransformBy(Matrix3d.Displacement(v));
                ent.TransformBy(Matrix3d.Displacement(-v));
                try { ent.RecordGraphicsModified(true); } catch { }
            }
            catch
            {
                // Best-effort only.
            }
        }
    }
}
