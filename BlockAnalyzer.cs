using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
        public static List<BlockInfo> AnalyzeFiles(List<string> filePaths)
        {
            System.Diagnostics.Debug.WriteLine($"[BlockAnalyzer] Rozpoczynam analizę {filePaths.Count} plików");
            var allBlocks = new Dictionary<string, BlockInfo>();

            foreach (var filePath in filePaths)
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine($"[BlockAnalyzer] Analizuję plik: {filePath}");
                    var blocks = AnalyzeFile(filePath);
                    System.Diagnostics.Debug.WriteLine($"[BlockAnalyzer] Znaleziono {blocks.Count} bloków w pliku {System.IO.Path.GetFileName(filePath)}");
                    
                    foreach (var block in blocks)
                    {
                        System.Diagnostics.Debug.WriteLine($"[BlockAnalyzer] Blok: {block.BlockName}, Atrybuty: {string.Join(", ", block.Attributes)}");
                        
                        if (!allBlocks.ContainsKey(block.BlockName))
                        {
                            allBlocks[block.BlockName] = new BlockInfo
                            {
                                BlockName = block.BlockName,
                                Attributes = new List<string>(block.Attributes)
                            };
                        }
                        else
                        {
                            // Merge attributes
                            foreach (var attr in block.Attributes)
                            {
                                if (!allBlocks[block.BlockName].Attributes.Contains(attr))
                                {
                                    allBlocks[block.BlockName].Attributes.Add(attr);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[BlockAnalyzer] Błąd analizy pliku {filePath}: {ex.Message}");
                    System.Diagnostics.Debug.WriteLine($"[BlockAnalyzer] StackTrace: {ex.StackTrace}");
                }
            }

            System.Diagnostics.Debug.WriteLine($"[BlockAnalyzer] Łącznie znaleziono {allBlocks.Count} unikalnych bloków");
            return allBlocks.Values.ToList();
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

        public static UpdateResult UpdateBlocksInFiles(List<string> filePaths, string blockName, string attributeName, string value, System.Action<int, int, string> progressCallback = null)
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
                    progressCallback?.Invoke(result.ProcessedFiles, result.TotalFiles, filePath);
                    
                    DwgVersion originalVersion = GetFileVersion(filePath);
                    
                    using (var db = new Database(false, true))
                    {
                        db.ReadDwgFile(filePath, FileOpenMode.OpenForReadAndAllShare, false, null);

                        using (var tr = db.TransactionManager.StartTransaction())
                        {
                            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                            bool updated = false;

                            // Update blocks in Model Space
                            var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                            updated |= UpdateBlocksInBlockTableRecord(tr, db, ms, blockName, attributeName, value);

                            // Update blocks in all layouts (Paper Space)
                            var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                            if (layoutDict != null)
                            {
                                foreach (DBDictionaryEntry entry in layoutDict)
                                {
                                    var layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                                    if (layout != null && layout.LayoutName != "Model")
                                    {
                                        var ps = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);
                                        updated |= UpdateBlocksInBlockTableRecord(tr, db, ps, blockName, attributeName, value);
                                    }
                                }
                            }

                            tr.Commit();
                            
                            if (updated)
                            {
                                db.SaveAs(filePath, originalVersion);
                            }
                            
                            result.SuccessfulFiles++;
                        }
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
                            if (dynamicBtr != null && dynamicBtr.Name.Equals(blockName, StringComparison.OrdinalIgnoreCase))
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
                                        
                                        // Restore formatting properties
                                        attRef.Justify = originalJustify;
                                        attRef.Height = originalHeight;
                                        attRef.WidthFactor = originalWidthFactor;
                                        attRef.Rotation = originalRotation;
                                        attRef.IsMirroredInX = originalIsMirroredInX;
                                        attRef.IsMirroredInY = originalIsMirroredInY;
                                        
                                        // Change the text - this is the only change we want
                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks] Changing TextString from '{originalText}' to '{value}'");
                                        attRef.TextString = value;
                                        
                                        // For centered text, we need to manually recalculate Position
                                        // because AutoCAD doesn't do it automatically when TextString changes
                                        if (hasAlignmentPoint)
                                        {
                                            // Store original WorkingDatabase and set it to current db
                                            Database oldDb = HostApplicationServices.WorkingDatabase;
                                            HostApplicationServices.WorkingDatabase = db;
                                            
                                            try
                                            {
                                                // Restore AlignmentPoint first
                                                attRef.AlignmentPoint = originalAlignmentPoint;
                                                
                                                // Use AdjustAlignment with WorkingDatabase set
                                                attRef.AdjustAlignment(db);
                                                System.Diagnostics.Debug.WriteLine($"[UpdateBlocks] Called AdjustAlignment with WorkingDatabase set");
                                                
                                                // Check if Position was recalculated
                                                var positionAfterAdjust = attRef.Position;
                                                var positionDelta = Math.Abs(positionAfterAdjust.X - originalPosition.X) + 
                                                                    Math.Abs(positionAfterAdjust.Y - originalPosition.Y);
                                                
                                                if (positionDelta < 0.0001)
                                                {
                                                    System.Diagnostics.Debug.WriteLine($"[UpdateBlocks] AdjustAlignment didn't recalculate Position, trying manual calculation");
                                                    
                                                    // Manual calculation: Use TextStyle to estimate text width
                                                    try
                                                    {
                                                        var textStyleId = attRef.TextStyleId;
                                                        var textStyle = tr.GetObject(textStyleId, OpenMode.ForRead) as TextStyleTableRecord;
                                                        
                                                        // Calculate approximate text width
                                                        // Width = char_count * char_width_factor * height * widthFactor
                                                        // For standard fonts, average char width is about 0.6-0.7 of height
                                                        double avgCharWidth = originalHeight * 0.65; // Approximate
                                                        double textWidth = (value?.Length ?? 0) * avgCharWidth * originalWidthFactor;
                                                        double halfWidth = textWidth / 2.0;
                                                        
                                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks] Estimated text width: {textWidth} (chars: {value?.Length ?? 0}, charWidth: {avgCharWidth}, halfWidth: {halfWidth})");
                                                        
                                                        // Calculate Position offset for BaseCenter
                                                        // Position = AlignmentPoint - (halfWidth) in text direction
                                                        double angleRad = originalRotation;
                                                        double deltaX = -halfWidth * Math.Cos(angleRad);
                                                        double deltaY = -halfWidth * Math.Sin(angleRad);
                                                        
                                                        Point3d calculatedPosition = new Point3d(
                                                            originalAlignmentPoint.X + deltaX,
                                                            originalAlignmentPoint.Y + deltaY,
                                                            originalAlignmentPoint.Z
                                                        );
                                                        
                                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks] Calculated Position: {calculatedPosition} (from AlignmentPoint {originalAlignmentPoint}, offset: {deltaX}, {deltaY})");
                                                        
                                                        attRef.Position = calculatedPosition;
                                                    }
                                                    catch (Exception ex2)
                                                    {
                                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks] Manual calculation failed: {ex2.Message}");
                                                    }
                                                }
                                                else
                                                {
                                                    System.Diagnostics.Debug.WriteLine($"[UpdateBlocks] Position was recalculated by AdjustAlignment: {positionAfterAdjust}");
                                                }
                                            }
                                            finally
                                            {
                                                HostApplicationServices.WorkingDatabase = oldDb;
                                            }
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
                                                        System.Diagnostics.Debug.WriteLine($"[UpdateBlocks]   Position should be recalculated but wasn't!");
                                                        
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
