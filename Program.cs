using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelGenerator {
    internal class Program {
        static async Task Main(string[] args) {
            var file = new FileInfo(args[0]);
            if (file.Exists) file.Delete();

            //CreateSpreadsheetWorkbook(args[0]);
            //InsertText(args[0], "test");
            var dt = GenerateData();
            var bytes = CreateDataTableExcelStream(dt);
            using (var fs = file.Create())
                await fs.WriteAsync(bytes, 0, bytes.Length);
        }

        private static DataTable GenerateData() {
            var dt = new DataTable();
            using (var conn = new SqlConnection(ConfigurationManager.ConnectionStrings["StockCrawler.Dao.Properties.Settings.StockConnectionString"].ConnectionString))
            using (var cmd = conn.CreateCommand()) {
                cmd.CommandText = "SELECT TOP 100 * FROM StockPriceHistory WITH(NOLOCK)";
                using (var da = new SqlDataAdapter(cmd))
                    da.Fill(dt);
            }

            return dt;
        }

        private static byte[] CreateDataTableExcelStream(DataTable dt) {
            using (var ms = new MemoryStream()) {
                using (var spreadSheet = CreateSpreadsheetWorkbook(ms)) {
                    for (var x = 0; x < dt.Columns.Count; x++)
                        InsertText(x, 1, dt.Columns[x].Caption, spreadSheet);

                    for (var y = 0; y < dt.Rows.Count; y++)
                        for (var x = 0; x < dt.Rows[y].ItemArray.Length; x++)
                            InsertText(x, (uint)(y + 2), dt.Rows[y][x].ToString(), spreadSheet);

                }
                return ms.ToArray();
            }
        }

        private static void InsertText(int x, uint y, string text, SpreadsheetDocument spreadSheet) {
            var workbookPart = spreadSheet.WorkbookPart ?? spreadSheet.AddWorkbookPart();

            // Get the SharedStringTablePart. If it does not exist, create a new one.
            SharedStringTablePart shareStringPart;
            if (workbookPart.GetPartsOfType<SharedStringTablePart>().Any())
                shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            else
                shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();

            // Insert the text into the SharedStringTablePart.
            var index = InsertSharedStringItem(text, shareStringPart);

            // Insert a new worksheet.
            var worksheetPart = InsertWorksheet(workbookPart);

            // Insert cell A1 into the new worksheet.
            var cell = InsertCellInWorksheet(GenerateColumnIndexString(x), y, worksheetPart);

            // Set the value of cell A1.
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            // Save the new worksheet.
            worksheetPart.Worksheet.Save();
        }
        private static string GenerateColumnIndexString(int x) {
            const char index_char = 'A';
            const int count_of_a_z = 26;
            var sb = new StringBuilder();
            var first_letter_index = x / count_of_a_z;
            char last_char = (char)(index_char + (x % count_of_a_z));
            for (var i = 0; i < first_letter_index; i++)
                sb.Append(index_char + i);

            sb.Append(last_char);
            return sb.ToString();
        }
        private static async void CreateSpreadsheetWorkbook(string filepath) {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            using (var ms = new MemoryStream()) {
                CreateSpreadsheetWorkbook(ms);
                using(var fs = File.Create(filepath)) 
                    await fs.WriteAsync(ms.ToArray(),0, (int)ms.Length);
            }
        }
        private static SpreadsheetDocument CreateSpreadsheetWorkbook(MemoryStream ms) {
            var spreadsheetDocument = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
            // Add a WorkbookPart to the document.
            var workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            var sheets = workbookPart.Workbook.AppendChild(new Sheets());

            //// Append a new worksheet and associate it with the workbook.
            //var sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
            //sheets.Append(sheet);

            workbookPart.Workbook.Save();
            return spreadsheetDocument;
        }

        // Given a document name and text, 
        // inserts a new work sheet and writes the text to cell "A1" of the new worksheet.
        private static void InsertText(string docName, string text) {
            // Open the document for editing.
            using (var spreadSheet = SpreadsheetDocument.Open(docName, true))
                InsertText(0, 1, text, spreadSheet);
        }
        // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart) {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable is null)
                shareStringPart.SharedStringTable = new SharedStringTable();

            var i = 0;

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
            WorksheetPart workSheetPart;
            if (!workbookPart.WorksheetParts.Any()) {
                // Add a new worksheet part to the workbook.
                workSheetPart = workbookPart.WorksheetParts.Any() ? workbookPart.WorksheetParts.First() : workbookPart.AddNewPart<WorksheetPart>();
                workSheetPart.Worksheet = new Worksheet(new SheetData());
                workSheetPart.Worksheet.Save();
            } else
                workSheetPart = workbookPart.WorksheetParts.First();

            var sheets = workbookPart.Workbook.GetFirstChild<Sheets>() ?? workbookPart.Workbook.AppendChild(new Sheets());
            var relationshipId = workbookPart.GetIdOfPart(workSheetPart);

            // Get a unique ID for the new sheet.
            uint sheetId = 1;
            if (!sheets.Elements<Sheet>().Any())
            {
                var sheetName = "Sheet" + sheetId;
                // Append the new worksheet and associate it with the workbook.
                var sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
                sheets.Append(sheet);
                workbookPart.Workbook.Save();
            }
            return workSheetPart;
        }
        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart) {
            var worksheet = worksheetPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();
            var cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;

            if (sheetData.Elements<Row>().Where(r => r.RowIndex != null && r.RowIndex == rowIndex).Any())
                row = sheetData.Elements<Row>().Where(r => r.RowIndex != null && r.RowIndex == rowIndex).First();
            else {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there != a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference != null && c.CellReference.Value == columnName + rowIndex).Any())
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