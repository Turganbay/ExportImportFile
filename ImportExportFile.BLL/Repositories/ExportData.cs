using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ImportExportFile.BLL.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;

namespace ImportExportFile.BLL.Repositories
{
    public class ExportData
    {
        Repository repo;
        public ExportData() 
        {
            repo = new Repository();
        }

        public void CreateSpreadsheetWorkbook(string filepath)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            FileInfo f = new FileInfo(filepath);
            //if (f.Exists)
                //f.Delete();

            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
            sheets.Append(sheet);

            // style
            WorkbookStylesPart stylesPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = GenerateStyleSheet();
            stylesPart.Stylesheet.Save();



            string cl = "";
            uint row = 2;
            int index;
            Cell cell;

            DataTable dt = new DataTable();
            dt.Columns.Add("Нефтепродукт", typeof(string));

            List<ExportList> exportData = repo.getExportData();

            List<string> r = new List<string>();
            List<string> p = new List<string>();

            Dictionary<string, string> dict = new Dictionary<string, string>();


            foreach (var item in exportData)
            {
                if (!r.Contains(item.region))
                {
                    r.Add(item.region);
                    dt.Columns.Add(item.region, typeof(string));

                }
                if (!p.Contains(item.product))
                {
                    p.Add(item.product);
                    //dt.Rows.Add(item.product);
                    dict[item.product] = "";
                }

                dict[item.product] += item.sum + "|";


            }

            foreach (string name in p)
            {
                int i = 0;
                DataRow newRow = dt.NewRow();
                string str = dict[name];
                char delimiterChar = '|';

                newRow[i] = name;
                string[] lines = str.Split(delimiterChar);

                foreach (string line in lines)
                {
                    i++;
                    if (!String.IsNullOrEmpty(line))
                    {
                        newRow[i] = line;
                    }
                }

                dt.Rows.Add(newRow);
            }



            foreach (DataRow dr in dt.Rows)
            {
                for (int idx = 0; idx < dt.Columns.Count; idx++)
                {
                    if (idx >= 26)
                        cl = "A" + Convert.ToString(Convert.ToChar(65 + idx - 26));
                    else
                        cl = Convert.ToString(Convert.ToChar(65 + idx));
                    SharedStringTablePart shareStringPart;
                    if (spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                    {
                        shareStringPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    }
                    else
                    {
                        shareStringPart = spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();
                    }
                    if (row == 2)
                    {
                        index = InsertSharedStringItem(dt.Columns[idx].ColumnName, shareStringPart);
                        cell = InsertCellInWorksheet(cl, row - 1, worksheetPart);
                        cell.CellValue = new CellValue(index.ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                        cell.StyleIndex = 1;
                    }
                    

                    // Insert the text into the SharedStringTablePart.
                    index = InsertSharedStringItem(Convert.ToString(dr[idx]), shareStringPart);
                    cell = InsertCellInWorksheet(cl, row, worksheetPart);
                    cell.CellValue = new CellValue(index.ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                    if (idx == 0)
                    {
                        cell.StyleIndex = 1;

                    }


                }
                row++;
            }
            workbookpart.Workbook.Save();

            // Close the document.
            spreadsheetDocument.Close();
            //InsertText(@"c:\MyXL3.xlx", "Hello");

        }

        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one. 
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {

                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                //foreach (Cell cell in row.Elements<Cell>())
                //{
                // if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                // {
                // refCell = cell;
                // break;
                // }
                //}
                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);
                worksheet.Save();
                return newCell;
            }
        }


        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }
            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }
                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();
            return i;
        }


        private Stylesheet GenerateStyleSheet()
        {
            return new Stylesheet(
                new Fonts(
                    new Font(                                                               // Index 0 – The default font.
                        new FontSize(){ Val = 11 },
                        new Color(){ Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName(){ Val = "Calibri" }),
                    new Font(                                                               // Index 1 – The bold font.
                        new Bold(),
                        new FontSize(){ Val = 11 },
                        new Color(){ Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName(){ Val = "Calibri" }),
                    new Font(                                                               // Index 2 – The Italic font.
                        new Italic(),
                        new FontSize(){ Val = 11 },
                        new Color(){ Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName(){ Val = "Calibri" }),
                    new Font(                                                               // Index 2 – The Times Roman font. with 16 size
                        new FontSize(){ Val = 16 },
                        new Color(){ Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName(){ Val = "Times New Roman" })
                ),
                new Fills(
                    new Fill(                                                           // Index 0 – The default fill.
                        new PatternFill(){ PatternType = PatternValues.None }),
                    new Fill(                                                           // Index 1 – The default fill of gray 125 (required)
                        new PatternFill(){ PatternType = PatternValues.DarkDown}),     
                    new Fill(                                                           // Index 2 – The yellow fill.
                        new PatternFill(
                            new ForegroundColor(){ Rgb = new HexBinaryValue() { Value = "FFFFFF00"} }
                        ){ PatternType = PatternValues.Solid })
                ),
                new Borders(
                    new Border(                                                         // Index 0 – The default border.
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder()),
                    new Border(                                                         // Index 1 – Applies a Left, Right, Top, Bottom border to a cell
                        new LeftBorder(
                            new Color(){ Auto = true }
                        ){ Style = BorderStyleValues.Thin },
                        new RightBorder(
                            new Color(){ Auto = true }
                        ){ Style = BorderStyleValues.Thin },
                        new TopBorder(
                            new Color(){ Auto = true }
                        ){ Style = BorderStyleValues.Thin },
                        new BottomBorder(
                            new Color(){ Auto = true }
                        ){ Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                ),
                new CellFormats(
                    new CellFormat(){ FontId = 0, FillId = 0, BorderId = 0},                          // Index 0 – The default cell style.  If a cell does not have a style index applied it will use this style combination instead
                    new CellFormat(){ FontId = 1, FillId = 2, BorderId = 1, ApplyFont = true },       // Index 1 – Bold 
                    new CellFormat(){ FontId = 2, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 2 – Italic
                    new CellFormat(){ FontId = 3, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 3 – Times Roman
                    new CellFormat(){ FontId = 0, FillId = 2, BorderId = 0, ApplyFill = true },       // Index 4 – Yellow Fill
                    new CellFormat(                                                                   // Index 5 – Alignment
                        new Alignment(){ Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
                    ){ FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true },
                    new CellFormat(){ FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }      // Index 6 – Border
                )
            ); // return
        }

    }
}