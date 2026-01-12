using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

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
                        var dm = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                        Document openedDoc = null;
                        Document previousDoc = null;
                        bool docWasAlreadyOpen = false;

                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocksInFiles] Opening document on UI thread: {filePath}");

                        previousDoc = dm.MdiActiveDocument;

                        // Detect if document is already open to avoid closing the user's drawing.
                        try
                        {
                            var normalizedFilePath = System.IO.Path.GetFullPath(filePath);
                            foreach (Document d in dm)
                            {
                                try
                                {
                                    var docPath = d.Database?.Filename;
                                    if (!string.IsNullOrEmpty(docPath) &&
                                        string.Equals(System.IO.Path.GetFullPath(docPath), normalizedFilePath, StringComparison.OrdinalIgnoreCase))
                                    {
                                        openedDoc = d;
                                        docWasAlreadyOpen = true;
                                        break;
                                    }
                                }
                                catch { }
                            }
                        }
                        catch { }

                        if (openedDoc == null)
                        {
                            openedDoc = dm.Open(filePath, false);
                        }
                        if (openedDoc == null)
                        {
                            fileError = "Failed to open document";
                            fileSucceeded = false;
                        }
                        else
                        {
                            try { dm.MdiActiveDocument = openedDoc; } catch { }

                            // 1) DB edits must be inside DocumentLock + Transaction
                            System.Diagnostics.Debug.WriteLine($"[UpdateBlocksInFiles] About to LockDocument (for DB edits): {filePath}");
                            using (openedDoc.LockDocument())
                            using (var tr = openedDoc.Database.TransactionManager.StartTransaction())
                            {
                                System.Diagnostics.Debug.WriteLine($"[UpdateBlocksInFiles] Locked + Transaction started: {filePath}");
                                var bt = (BlockTable)tr.GetObject(openedDoc.Database.BlockTableId, OpenMode.ForRead);
                                updatedInFile = false;

                                // Model space
                                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                                updatedInFile |= UpdateBlocksInBlockTableRecord(tr, openedDoc.Database, ms, blockName, attributeName, value);

                                // Paper space layouts
                                var layoutDict = (DBDictionary)tr.GetObject(openedDoc.Database.LayoutDictionaryId, OpenMode.ForRead);
                                if (layoutDict != null)
                                {
                                    foreach (DBDictionaryEntry entry in layoutDict)
                                    {
                                        var layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                                        if (layout != null && layout.LayoutName != "Model")
                                        {
                                            var ps = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);
                                            updatedInFile |= UpdateBlocksInBlockTableRecord(tr, openedDoc.Database, ps, blockName, attributeName, value);
                                        }
                                    }
                                }

                                tr.Commit();
                            } // release DocumentLock BEFORE running commands

                            if (updatedInFile)
                            {
                                // We already re-sync the modified AttributeReference from AttributeDefinition using
                                // SetAttributeFromBlock + AdjustAlignment inside UpdateBlocksInBlockTableRecord.
                                // That is the API equivalent of ATTSYNC for the changed attributes, without relying
                                // on the command system (which is unstable here).
                                try { openedDoc.Editor.Regen(); } catch { }

                                // Save reliably:
                                // - If we opened the doc, CloseAndSave uses AutoCAD's save pipeline (stable).
                                // - If doc was already open, we do NOT close it; saving automatically is risky here,
                                //   so we leave it open and report success (changes are in-memory).
                                if (!docWasAlreadyOpen)
                                {
                                    try
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocksInFiles] CloseAndSave: {filePath}");
                                        openedDoc.CloseAndSave(filePath);
                                        openedDoc = null; // it's closed now
                                        fileSucceeded = true; // IMPORTANT: mark success after successful save
                                    }
                                    catch (Exception saveEx)
                                    {
                                        fileError = $"Save failed: {saveEx.GetType().Name}: {saveEx.Message}";
                                        fileSucceeded = false;
                                    }
                                }
                                else
                                {
                                    // Best-effort: leave document open; user can save normally.
                                    System.Diagnostics.Debug.WriteLine($"[UpdateBlocksInFiles] Document already open; leaving it open for user to save: {filePath}");
                                    fileSucceeded = true;
                                }
                            }
                            else
                            {
                                fileSucceeded = true; // nothing to change, but no error
                            }

                            // Close only if we opened it in this run AND didn't already CloseAndSave.
                            try
                            {
                                if (openedDoc != null && !docWasAlreadyOpen)
                                {
                                    openedDoc.CloseAndDiscard();
                                }
                            }
                            catch { }

                            try
                            {
                                if (previousDoc != null)
                                {
                                    dm.MdiActiveDocument = previousDoc;
                                }
                            }
                            catch { }
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
                                        
                                        // Immediately after text change, check what changed
                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks] AFTER TEXT CHANGE:");
                                        var afterPosition = attRef.Position;
                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   Position={afterPosition} (was {originalPosition})");
                                        
                                        if (hasAlignmentPoint)
                                        {
                                            try
                                            {
                                                var afterAlignmentPoint = attRef.AlignmentPoint;
                                                System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   AlignmentPoint={afterAlignmentPoint} (was {originalAlignmentPoint})");
                                                
                                                // Check if AlignmentPoint changed
                                                var deltaX = Math.Abs(afterAlignmentPoint.X - originalAlignmentPoint.X);
                                                var deltaY = Math.Abs(afterAlignmentPoint.Y - originalAlignmentPoint.Y);
                                                var deltaZ = Math.Abs(afterAlignmentPoint.Z - originalAlignmentPoint.Z);
                                                
                                                if (deltaX > 0.0001 || deltaY > 0.0001 || deltaZ > 0.0001)
                                                {
                                                    System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   AlignmentPoint CHANGED! Delta: X={deltaX}, Y={deltaY}, Z={deltaZ}");
                                                    attRef.AlignmentPoint = originalAlignmentPoint;
                                                    System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   Restored AlignmentPoint to {originalAlignmentPoint}");
                                                    
                                                    // Check Position after restoring AlignmentPoint
                                                    var finalPosition = attRef.Position;
                                                    System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   Position after restore={finalPosition}");
                                                }
                                                else
                                                {
                                                    System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   AlignmentPoint unchanged");
                                                    
                                                    // Even if AlignmentPoint is unchanged, Position might need recalculation
                                                    // Calculate expected Position based on text length difference
                                                    var textLengthDiff = (value?.Length ?? 0) - originalText.Length;
                                                    if (textLengthDiff != 0)
                                                    {
                                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   Text length changed by {textLengthDiff} characters");
                                                        
                                                        // If Position already changed, the re-sync did its job.
                                                        var posDx = Math.Abs(afterPosition.X - originalPosition.X);
                                                        var posDy = Math.Abs(afterPosition.Y - originalPosition.Y);
                                                        var posDz = Math.Abs(afterPosition.Z - originalPosition.Z);
                                                        var positionChanged = posDx > 0.0001 || posDy > 0.0001 || posDz > 0.0001;
                                                        
                                                        if (positionChanged)
                                                        {
                                                            System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   Position was recalculated (delta: X={posDx}, Y={posDy}, Z={posDz})");
                                                        }
                                                        else
                                                        {
                                                            System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   WARNING: Position did not change even though text length changed");
                                                        }
                                                        
                                                        // Try to manually recalculate Position
                                                        // For BaseCenter, Position = AlignmentPoint - (text_width / 2)
                                                        // We can't easily calculate text width, so try AdjustAlignment again
                                                        try
                                                        {
                                                            attRef.AdjustAlignment(db);
                                                            var recalcPosition = attRef.Position;
                                                            System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   Position after AdjustAlignment={recalcPosition}");
                                                        }
                                                        catch
                                                        {
                                                            System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   AdjustAlignment failed again");
                                                        }
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   ERROR reading AlignmentPoint: {ex.Message}");
                                            }
                                        }
                                        else
                                        {
                                            // For left-aligned text, check if Position changed
                                            var deltaX = Math.Abs(afterPosition.X - originalPosition.X);
                                            var deltaY = Math.Abs(afterPosition.Y - originalPosition.Y);
                                            var deltaZ = Math.Abs(afterPosition.Z - originalPosition.Z);
                                            
                                            if (deltaX > 0.0001 || deltaY > 0.0001 || deltaZ > 0.0001)
                                            {
                                                System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   Position CHANGED! Delta: X={deltaX}, Y={deltaY}, Z={deltaZ}");
                                                attRef.Position = originalPosition;
                                                System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   Restored Position to {originalPosition}");
                                            }
                                            else
                                            {
                                                System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   Position unchanged");
                                            }
                                        }
                                        
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
    }
}
