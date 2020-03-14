using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Reflection;

namespace DelegationGeneratorConsoleApp
{
    class Program
    {
        const string oldFile = "DelegationPattern.pdf";
        const string newFile = "delegation0{0}.pdf";
        const string excelFileName = "DelegationData.xlsx";

        static void Main()
        {
            using (var dt = DataTableFromExcelFile(excelFileName))
                GenerateDelegationDocuments(dt);
        }

        private static void GenerateDelegationDocuments(DataTable dt)
        {
            for (var i = 0; i < dt.Rows.Count; i++)
            {
                var dateText = (string)dt.Rows[i][0];
                var dateDateTime = DateTime.ParseExact(dateText, "dd.MM.yyyy", CultureInfo.InvariantCulture);
                var timeStart1 = (string)dt.Rows[i][1];
                var timeStart2 = (string)dt.Rows[i][2];
                var timeEnd1 = (string)dt.Rows[i][3];
                var timeEnd2 = (string)dt.Rows[i][4];
                var tailNumber = (i + 1).ToString("00");
                var currentFile = string.Format(newFile, tailNumber);
                var delegationNumber = tailNumber + "/" + dateDateTime.ToString("MM") + "/2019";
                var reader = new PdfReader(oldFile);
                var size = reader.GetPageSizeWithRotation(1);
                var document = new Document(size);
                var fs = new FileStream(currentFile, FileMode.Create, FileAccess.Write);
                var writer = PdfWriter.GetInstance(document, fs);
                document.Open();
                document.NewPage();
                PdfContentByte cb = writer.DirectContent;
                var bf = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetColorFill(BaseColor.BLACK);
                cb.SetFontAndSize(bf, 12);
                //Business travel no.
                cb.BeginText();
                cb.ShowTextAligned(1, delegationNumber, 244, 652, 0);
                cb.EndText();
                //Business travel date
                cb.BeginText();
                cb.ShowTextAligned(1, dateText, 175, 625, 0);
                cb.EndText();
                //Business travel date from
                cb.BeginText();
                cb.ShowTextAligned(1, dateText, 155, 460, 0);
                cb.EndText();
                //Business travel date from
                cb.BeginText();
                cb.ShowTextAligned(1, dateText, 240, 460, 0);
                cb.EndText();
                //Signature date
                cb.BeginText();
                cb.ShowTextAligned(1, dateText, 121, 317, 0);
                cb.EndText();
                //Business travel expenditures
                cb.BeginText();
                cb.ShowTextAligned(1, delegationNumber, 380, 226, 0);
                cb.EndText();
                //Check date
                cb.BeginText();
                cb.ShowTextAligned(1, dateText, 320, 76, 0);
                cb.EndText();
                PdfImportedPage page1 = writer.GetImportedPage(reader, 1);
                cb.AddTemplate(page1, 0, 0);
                document.NewPage();
                bf = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetColorFill(BaseColor.BLACK);
                cb.SetFontAndSize(bf, 8);
                //Business travel date start1
                cb.BeginText();
                cb.ShowTextAligned(1, dateText, 165, 733, 0);
                cb.EndText();
                //Business travel date end1
                cb.BeginText();
                cb.ShowTextAligned(1, dateText, 165, 719, 0);
                cb.EndText();
                //Business travel date start2
                cb.BeginText();
                cb.ShowTextAligned(1, dateText, 320, 733, 0);
                cb.EndText();
                //Business travel date end2
                cb.BeginText();
                cb.ShowTextAligned(1, dateText, 320, 719, 0);
                cb.EndText();
                //Business travel time start1
                cb.BeginText();
                cb.ShowTextAligned(1, timeStart1, 210, 733, 0);
                cb.EndText();
                //Business travel time end1
                cb.BeginText();
                cb.ShowTextAligned(1, timeEnd1, 210, 719, 0);
                cb.EndText();
                //Business travel time start2
                cb.BeginText();
                cb.ShowTextAligned(1, timeStart2, 365, 733, 0);
                cb.EndText();
                //Business travel time end2
                cb.BeginText();
                cb.ShowTextAligned(1, timeEnd2, 365, 719, 0);
                cb.EndText();
                //Formal check data
                cb.BeginText();
                cb.ShowTextAligned(1, dateText, 92, 530, 0);
                cb.EndText();
                //Meritorical check data
                cb.BeginText();
                cb.ShowTextAligned(1, dateText, 217, 530, 0);
                cb.EndText();
                bf = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb.SetColorFill(BaseColor.BLACK);
                cb.SetFontAndSize(bf, 12);
                //Confirmation data
                cb.BeginText();
                cb.ShowTextAligned(1, dateText, 115, 418, 0);
                cb.EndText();
                //Receipt data
                cb.BeginText();
                cb.ShowTextAligned(1, dateText, 100, 316, 0);
                cb.EndText();
                //Bill data
                cb.BeginText();
                cb.ShowTextAligned(1, dateText, 350, 326, 0);
                cb.EndText();
                //Delegate data signature
                cb.BeginText();
                cb.ShowTextAligned(1, dateText, 384, 188, 0);
                cb.EndText();
                PdfImportedPage page2 = writer.GetImportedPage(reader, 2);
                cb.AddTemplate(page2, 0, 0);
                document.Close();
                fs.Close();
                writer.Close();
                reader.Close();
            }
        }

        private static DataTable DataTableFromExcelFile(string excelFileName)
        {
            var path = 
                Path.GetDirectoryName(
                    Assembly.GetExecutingAssembly().GetName().CodeBase
                );
            path = Path.Combine(path, excelFileName);
            var xlApp = new Excel.Application();
            var xlWorkbook = xlApp.Workbooks.Open(path);
            var xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];
            var dt = GetWorksheetAsDataTable(xlWorksheet);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.FinalReleaseComObject(xlWorksheet);
            xlWorkbook.Close();
            Marshal.FinalReleaseComObject(xlWorkbook);
            xlApp.Quit();
            Marshal.FinalReleaseComObject(xlApp);
            return dt;
        }

        public static DataTable GetWorksheetAsDataTable(Excel.Worksheet worksheet)
        {
            var dt = new DataTable(worksheet.Name);
            dt.Columns.AddRange(
                GatherColumnNames(worksheet).Select(
                    x => new DataColumn(x)
                ).ToArray()
            );
            var headerOffset = 1;
            var width = dt.Columns.Count;
            var depth = GetTableDepth(worksheet, headerOffset);
            for (var i = 1; i <= depth; i++)
            {
                var row = dt.NewRow();
                for (var j = 1; j <= width; j++)
                {
                    var currentValue = worksheet.Cells[i + headerOffset, j].Value;
                    row[j - 1] = currentValue?.ToString();
                }
                dt.Rows.Add(row);
            }
            return dt;
        }

        private static int GetTableDepth(Excel.Worksheet worksheet, int headerOffset)
        {
            var i = 1;
            var j = 1;
            var cellValue = worksheet.Cells[i + headerOffset, j].Value;
            while (cellValue != null)
            {
                i++;
                cellValue = worksheet.Cells[i + headerOffset, j].Value;
            }
            return i - 1;
        }

        private static IEnumerable<string> GatherColumnNames(Excel.Worksheet worksheet)
        {
            var columns = new List<string>();
            var i = 1;
            var j = 1;
            var columnName = worksheet.Cells[i, j].Value;
            while (columnName != null)
            {
                columns.Add(GetUniqueColumnName(columns, columnName.ToString()));
                j++;
                columnName = worksheet.Cells[i, j].Value;
            }
            return columns;
        }

        private static string GetUniqueColumnName(IEnumerable<string> columnNames, string columnName)
        {
            var colName = columnName;
            var i = 1;
            while (columnNames.Contains(colName))
            {
                colName = columnName + i.ToString();
                i++;
            }
            return colName;
        }
    }
}
