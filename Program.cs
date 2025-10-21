using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelMonitoring1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Batam Schneider Etry Data (Batam S.E.D) Program");

            Console.Write("Enter file path (or press Enter for default): ");
            string? inputPath = Console.ReadLine();
            string filePath = string.IsNullOrWhiteSpace(inputPath)
                ? "Batam_Schneider_Etry_Data_File1.xlsx"
                : inputPath.Trim();

            Console.Write("Please, input your Data Name (sheet name): ");
            string? inputSheet = Console.ReadLine();
            string sheetName = string.IsNullOrWhiteSpace(inputSheet) ? "Data1" : inputSheet.Trim();

            EnsureWorkbookAndSheet(filePath, sheetName);

            while (true)
            {
                Console.WriteLine();
                Console.WriteLine("Choose: (C)reate row, (R)ead rows, (U)pdate row, (D)elete row, (Q)uit");
                Console.Write("Selection: ");
                string? sel = Console.ReadLine();
                if (string.IsNullOrWhiteSpace(sel)) continue;
                sel = sel.Trim().ToUpperInvariant();

                try
                {
                    if (sel == "C")
                    {
                        Console.Write("Enter comma-separated values to add: ");
                        var values = ReadCsvLine();
                        AddRow(filePath, sheetName, values);
                        Console.WriteLine("Row added.");
                    }
                    else if (sel == "R")
                    {
                        var rows = ReadRows(filePath, sheetName);
                        if (!rows.Any()) Console.WriteLine("(no rows)");
                        else
                        {
                            int i = 1;
                            foreach (var r in rows)
                            {
                                Console.WriteLine($"{i++}: {string.Join(", ", r)}");
                            }
                        }
                    }
                    else if (sel == "U")
                    {
                        Console.Write("Enter row index to update (1-based): ");
                        var idxInput = Console.ReadLine();
                        if (!int.TryParse(idxInput, out int idx) || idx < 1) { Console.WriteLine("Invalid index."); continue; }
                        Console.Write("Enter comma-separated new values: ");
                        var values = ReadCsvLine();
                        bool ok = UpdateRow(filePath, sheetName, idx, values);
                        Console.WriteLine(ok ? "Row updated." : "Row not found.");
                    }
                    else if (sel == "D")
                    {
                        Console.Write("Enter row index to delete (1-based): ");
                        var idxInput = Console.ReadLine();
                        if (!int.TryParse(idxInput, out int idx) || idx < 1) { Console.WriteLine("Invalid index."); continue; }
                        bool ok = DeleteRow(filePath, sheetName, idx);
                        Console.WriteLine(ok ? "Row deleted." : "Row not found.");
                    }
                    else if (sel == "Q")
                    {
                        break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                }
            }
        }

        static string[] ReadCsvLine()
        {
            var line = Console.ReadLine() ?? string.Empty;
            return line.Split(',').Select(s => s.Trim()).ToArray();
        }

        static void EnsureWorkbookAndSheet(string path, string sheetName)
        {
            // Create file if missing
            if (!File.Exists(path))
            {
                using var createDoc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
                var createWbPart = createDoc.AddWorkbookPart();
                createWbPart.Workbook = new Workbook();
                var worksheetPart = createWbPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                var sheets = createWbPart.Workbook.AppendChild(new Sheets());
                var sheet = new Sheet()
                {
                    Id = createWbPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = sheetName
                };
                sheets.Append(sheet);
                createWbPart.Workbook.Save();
                return;
            }

            // Open existing and ensure sheet exists
            using var doc = SpreadsheetDocument.Open(path, true);
            var wbPart = doc.WorkbookPart ?? doc.AddWorkbookPart();
            wbPart.Workbook ??= new Workbook();

            var sheetsElement = wbPart.Workbook.GetFirstChild<Sheets>() ?? wbPart.Workbook.AppendChild(new Sheets());

            var existing = sheetsElement.Elements<Sheet>().FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
            if (existing == null)
            {
                var worksheetPart = wbPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                uint newId = sheetsElement.Elements<Sheet>().Select(s => s.SheetId?.Value ?? 0u).DefaultIfEmpty(0u).Max() + 1;
                var sheet = new Sheet()
                {
                    Id = wbPart.GetIdOfPart(worksheetPart),
                    SheetId = newId,
                    Name = sheetName
                };
                sheetsElement.Append(sheet);
                wbPart.Workbook.Save();
            }
        }

        static WorksheetPart? GetWorksheetPartByName(WorkbookPart workbookPart, string sheetName)
        {
            var sheets = workbookPart.Workbook?.GetFirstChild<Sheets>();
            if (sheets == null) return null;
            var sheet = sheets.Elements<Sheet>().FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
            if (sheet == null) return null;
            var id = sheet.Id?.Value;
            if (string.IsNullOrEmpty(id)) return null;
            return workbookPart.GetPartById(id) as WorksheetPart;
        }

        static List<string[]> ReadRows(string path, string sheetName)
        {
            var result = new List<string[]>();
            using var doc = SpreadsheetDocument.Open(path, false);
            var wbPart = doc.WorkbookPart;
            if (wbPart == null) return result;
            var wsPart = GetWorksheetPartByName(wbPart, sheetName);
            if (wsPart == null) return result;
            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return result;

            var sst = wbPart.SharedStringTablePart?.SharedStringTable;

            foreach (var row in sheetData.Elements<Row>())
            {
                var cells = row.Elements<Cell>().ToArray();
                var values = new List<string>();
                foreach (var cell in cells)
                {
                    string? cellText = GetCellText(cell, sst);
                    values.Add(cellText ?? string.Empty);
                }
                result.Add(values.ToArray());
            }
            return result;
        }

        static string? GetCellText(Cell cell, SharedStringTable? sst)
        {
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                if (int.TryParse(cell.CellValue?.InnerText, out int sstIndex) && sst != null)
                {
                    var ssi = sst.Elements<SharedStringItem>().ElementAtOrDefault(sstIndex);
                    return ssi?.InnerText ?? ssi?.Text?.Text;
                }
                return cell.CellValue?.InnerText;
            }
            return cell.CellValue?.InnerText;
        }

        static int InsertSharedStringItem(WorkbookPart wbPart, string text)
        {
            var sstPart = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault() ?? wbPart.AddNewPart<SharedStringTablePart>();
            var sst = sstPart.SharedStringTable ?? new SharedStringTable();

            // Try find existing
            int index = 0;
            foreach (var item in sst.Elements<SharedStringItem>())
            {
                if ((item.Text?.Text ?? item.InnerText) == text)
                    return index;
                index++;
            }

            sst.AppendChild(new SharedStringItem(new Text(text)));
            sstPart.SharedStringTable = sst;
            sstPart.SharedStringTable.Save();
            return index;
        }

        static void AddRow(string path, string sheetName, string[] values)
        {
            using var doc = SpreadsheetDocument.Open(path, true);
            var wbPart = doc.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart missing.");
            var wsPart = GetWorksheetPartByName(wbPart, sheetName) ?? throw new InvalidOperationException("Worksheet not found.");
            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>() ?? wsPart.Worksheet.AppendChild(new SheetData());

            uint nextRowIndex = 1;
            var lastRow = sheetData.Elements<Row>().LastOrDefault();
            if (lastRow != null) nextRowIndex = (lastRow.RowIndex?.Value ?? 0u) + 1u;

            var row = new Row() { RowIndex = nextRowIndex };
            for (int i = 0; i < values.Length; i++)
            {
                var text = values[i] ?? string.Empty;
                int sstIndex = InsertSharedStringItem(wbPart, text);

                var cell = new Cell()
                {
                    CellReference = GetCellReference(i, (int)nextRowIndex),
                    DataType = new EnumValue<CellValues>(CellValues.SharedString),
                    CellValue = new CellValue(sstIndex.ToString())
                };
                row.Append(cell);
            }
            sheetData.Append(row);
            wsPart.Worksheet.Save();
            wbPart.Workbook.Save();
        }

        static bool UpdateRow(string path, string sheetName, int rowIndex, string[] values)
        {
            using var doc = SpreadsheetDocument.Open(path, true);
            var wbPart = doc.WorkbookPart;
            if (wbPart == null) return false;
            var wsPart = GetWorksheetPartByName(wbPart, sheetName);
            if (wsPart == null) return false;
            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return false;

            var row = sheetData.Elements<Row>().FirstOrDefault(r => (r.RowIndex?.Value ?? 0u) == (uint)rowIndex);
            if (row == null) return false;

            row.RemoveAllChildren<Cell>();
            for (int i = 0; i < values.Length; i++)
            {
                var text = values[i] ?? string.Empty;
                int sstIndex = InsertSharedStringItem(wbPart, text);

                var cell = new Cell()
                {
                    CellReference = GetCellReference(i, rowIndex),
                    DataType = new EnumValue<CellValues>(CellValues.SharedString),
                    CellValue = new CellValue(sstIndex.ToString())
                };
                row.Append(cell);
            }
            wsPart.Worksheet.Save();
            wbPart.Workbook.Save();
            return true;
        }

        static bool DeleteRow(string path, string sheetName, int rowIndex)
        {
            using var doc = SpreadsheetDocument.Open(path, true);
            var wbPart = doc.WorkbookPart;
            if (wbPart == null) return false;
            var wsPart = GetWorksheetPartByName(wbPart, sheetName);
            if (wsPart == null) return false;
            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return false;

            var row = sheetData.Elements<Row>().FirstOrDefault(r => (r.RowIndex?.Value ?? 0u) == (uint)rowIndex);
            if (row == null) return false;

            row.Remove();
            wsPart.Worksheet.Save();
            wbPart.Workbook.Save();
            return true;
        }

        static string GetCellReference(int columnIndexZeroBased, int rowIndex)
        {
            int dividend = columnIndexZeroBased + 1;
            string columnName = String.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return $"{columnName}{rowIndex}";
        }

        record EntryRecord(
            string Date,          // "yyyy-MM-dd" preferred
            string Shift,
            string CodeReference,
            string MachineNumber,
            string Area,          // e.g. "Backend 1"
            string ProcessAutoAdjustment,
            string ProcessTopTec,
            string ProcessFinalTester,
            string ProcessPackaging,
            int QuantityInput,
            int QuantityGood,
            int QuantityBad,
            int Reject
        )
        {
            public string[] ToValues() =>
                new string[]
                {
                    Date,
                    Shift,
                    CodeReference,
                    MachineNumber,
                    Area,
                    ProcessAutoAdjustment,
                    ProcessTopTec,
                    ProcessFinalTester,
                    ProcessPackaging,
                    QuantityInput.ToString(),
                    QuantityGood.ToString(),
                    QuantityBad.ToString(),
                    Reject.ToString()
                };

            public static string[] Headers => new[]
            {
                "Date","Shift","CodeReference","MachineNumber","Area",
                "AutoAdjustment","TopTec","FinalTester","Packaging",
                "QuantityInput","QuantityGood","QuantityBad","Reject"
            };
        }

        // Helper: prompt user for a typed record (place inside Program class)
        static EntryRecord PromptForEntryRecord()
        {
            Console.WriteLine("Enter record values (press Enter to skip / default empty):");
            string Read(string prompt)
            {
                Console.Write(prompt);
                return (Console.ReadLine() ?? string.Empty).Trim();
            }

            string date = Read("Date (yyyy-MM-dd): ");
            string shift = Read("Shift: ");
            string code = Read("Code Reference: ");
            string machine = Read("Machine Number: ");
            string area = Read("Area (Backend 1/2/3): ");
            string autoAdj = Read("Process Auto Adjustment: ");
            string topTec = Read("Process TopTec: ");
            string final = Read("Process Final Tester: ");
            string pack = Read("Process Packaging: ");

            static int ReadInt(string label)
            {
                Console.Write(label);
                var s = Console.ReadLine();
                return int.TryParse(s, out var v) ? v : 0;
            }

            int qtyIn = ReadInt("Quantity Input: ");
            int qtyGood = ReadInt("Quantity Good: ");
            int qtyBad = ReadInt("Quantity Bad: ");
            int reject = ReadInt("Reject: ");

            return new EntryRecord(date, shift, code, machine, area, autoAdj, topTec, final, pack, qtyIn, qtyGood, qtyBad, reject);
        }

        // Ensure header row exists for the sheet (call after EnsureWorkbookAndSheet)
        static void EnsureHeaderRow(string path, string sheetName)
        {
            using var doc = SpreadsheetDocument.Open(path, true);
            var wbPart = doc.WorkbookPart ?? throw new InvalidOperationException("Workbook missing");
            var wsPart = GetWorksheetPartByName(wbPart, sheetName) ?? throw new InvalidOperationException("Worksheet not found");
            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>() ?? wsPart.Worksheet.AppendChild(new SheetData());

            var firstRow = sheetData.Elements<Row>().FirstOrDefault();
            var headers = EntryRecord.Headers;
            bool needsHeader = true;

            if (firstRow != null)
            {
                var firstValues = firstRow.Elements<Cell>().Select(c => GetCellText(c, wbPart.SharedStringTablePart?.SharedStringTable) ?? string.Empty).ToArray();
                if (firstValues.Length >= headers.Length && headers.SequenceEqual(firstValues.Take(headers.Length)))
                    needsHeader = false;
            }

            if (needsHeader)
            {
                var headerRow = new Row() { RowIndex = 1u };
                for (int i = 0; i < headers.Length; i++)
                {
                    int sstIndex = InsertSharedStringItem(wbPart, headers[i]);
                    var cell = new Cell()
                    {
                        CellReference = GetCellReference(i, 1),
                        DataType = new EnumValue<CellValues>(CellValues.SharedString),
                        CellValue = new CellValue(sstIndex.ToString())
                    };
                    headerRow.Append(cell);
                }
                // Shift existing rows down if needed (simple append if empty)
                sheetData.InsertAt(headerRow, 0);
                wsPart.Worksheet.Save();
                wbPart.Workbook.Save();
            }
        }

        // Convenience: add typed record (uses existing InsertSharedStringItem)
        static void AddEntryRecord(string path, string sheetName, EntryRecord rec)
        {
            AddRow(path, sheetName, rec.ToValues());
        }
    }
}