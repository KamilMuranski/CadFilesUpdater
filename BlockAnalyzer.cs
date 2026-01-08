using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
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

        public static bool UpdateBlocksInFiles(List<string> filePaths, string blockName, string attributeName, string value)
        {
            bool allSuccess = true;

            foreach (var filePath in filePaths)
            {
                try
                {
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
                            updated |= UpdateBlocksInBlockTableRecord(tr, ms, blockName, attributeName, value);

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
                                        updated |= UpdateBlocksInBlockTableRecord(tr, ps, blockName, attributeName, value);
                                    }
                                }
                            }

                            tr.Commit();
                            
                            if (updated)
                            {
                                db.SaveAs(filePath, originalVersion);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Błąd aktualizacji pliku {filePath}: {ex.Message}");
                    allSuccess = false;
                }
            }

            return allSuccess;
        }

        private static bool UpdateBlocksInBlockTableRecord(Transaction tr, BlockTableRecord btr, string blockName, string attributeName, string value)
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
                                        attRef.TextString = value;
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
