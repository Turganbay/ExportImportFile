﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ImportExportFile.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;

namespace ImportExportFile.Repository
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
            if (f.Exists)
                f.Delete();

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
                    }

                    // Insert the text into the SharedStringTablePart.
                    index = InsertSharedStringItem(Convert.ToString(dr[idx]), shareStringPart);
                    cell = InsertCellInWorksheet(cl, row, worksheetPart);
                    cell.CellValue = new CellValue(index.ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
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

    }
}