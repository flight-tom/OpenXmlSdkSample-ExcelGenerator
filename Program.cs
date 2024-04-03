using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Linq;

namespace ExcelGenerator {
    internal class Program {
        static void Main(string[] args) {
            CreateSpreadsheetWorkbook(args[0]);
            InsertText(args[0], "test");
        }
        private static void CreateSpreadsheetWorkbook(string filepath) {
            var file = new FileInfo(filepath);
            if (file.Exists) file.Delete();

            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            using (var spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook)) {
                // Add a WorkbookPart to the document.
                var workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add Sheets to the Workbook.
                var sheets = workbookPart.Workbook.AppendChild(new Sheets());

                // Append a new worksheet and associate it with the workbook.
                var sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
                sheets.Append(sheet);

                workbookPart.Workbook.Save();
            }
        }
        // Given a document name and text, 
        // inserts a new work sheet and writes the text to cell "A1" of the new worksheet.
        private static void InsertText(string docName, string text) {
            // Open the document for editing.
            using (var spreadSheet = SpreadsheetDocument.Open(docName, true)) {
                var workbookPart = spreadSheet.WorkbookPart ?? spreadSheet.AddWorkbookPart();

                // Get the SharedStringTablePart. If it does not exist, create a new one.
                SharedStringTablePart shareStringPart;
                if (workbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                    shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                else
                    shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();

                // Insert the text into the SharedStringTablePart.
                var index = InsertSharedStringItem(text, shareStringPart);

                // Insert a new worksheet.
                var worksheetPart = InsertWorksheet(workbookPart);

                // Insert cell A1 into the new worksheet.
                var cell = InsertCellInWorksheet("A", 1, worksheetPart);

                // Set the value of cell A1.
                cell.CellValue = new CellValue(index.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                // Save the new worksheet.
                worksheetPart.Worksheet.Save();
            }
        }
        // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart) {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable is null)
                shareStringPart.SharedStringTable = new SharedStringTable();

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (var item in shareStringPart.SharedStringTable.Elements<SharedStringItem>()) {
                if (item.InnerText == text)
                    return i;

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }
        // Given a WorkbookPart, inserts a new worksheet.
        private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart) {
            // Add a new worksheet part to the workbook.
            var newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            var sheets = workbookPart.Workbook.GetFirstChild<Sheets>() ?? workbookPart.Workbook.AppendChild(new Sheets());
            var relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new sheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
                sheetId = sheets.Elements<Sheet>().Select<Sheet, uint>(s => {
                    if (s.SheetId != null && s.SheetId.HasValue)
                        return s.SheetId.Value;

                    return 0;
                }).Max() + 1;

            var sheetName = "Sheet" + sheetId;

            // Append the new worksheet and associate it with the workbook.
            var sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }
        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart) {
            var worksheet = worksheetPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();
            var cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;

            if (sheetData.Elements<Row>().Where(r => r.RowIndex != null && r.RowIndex == rowIndex).Count() != 0)
                row = sheetData.Elements<Row>().Where(r => r.RowIndex != null && r.RowIndex == rowIndex).First();
            else {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there != a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference != null && c.CellReference.Value == columnName + rowIndex).Count() > 0)
                return row.Elements<Cell>().Where(c => c.CellReference != null && c.CellReference.Value == cellReference).First();
            else {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;

                foreach (var cell in row.Elements<Cell>())
                    if (string.Compare(cell.CellReference?.Value, cellReference, true) > 0) {
                        refCell = cell;
                        break;
                    }

                var newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }
    }
}