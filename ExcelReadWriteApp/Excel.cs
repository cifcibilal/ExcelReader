using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAcmaVeOkumaEgitim
{
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        public Worksheet ws;
        public Excel() { }
        public Excel(string path, int Sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
        }
        public string ReadCell(int i, int j)
        {
            i++; j++;
            if (ws.Cells[i, j].Value2 != null)
            {
                return ws.Cells[i, j].Value2;
            }
            else { return ""; }
        }
        public void WriteToCell(int i, int j, string s)
        {
            i++;
            j++;
            ws.Cells[i, j].Value2 = s;
        }
        public void Save()
        {
            wb.Save();
        }
        public void SaveAs(string path)
        {
            wb.SaveAs(path);
        }
        public void Close()
        {
            wb.Close();
        }
        public void CreateNewFile()
        {
            this.wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            this.ws = wb.Worksheets[1];
        }
        public void CreateNewSheet()
        {
            Worksheet temptSheet = wb.Worksheets.Add(After: ws);
        }
        public void SelectWorkSheet(int SheetNumber)
        {
            this.ws = wb.Worksheets[SheetNumber];
        }
        public void DeleteWorkSheet(int SheetNumber)
        {
            wb.Worksheets[SheetNumber].Delete();
        }
        public void ProtectSheet()
        {
            ws.Protect();
        }
        public void ProtectSheet(string Password)
        {
            ws.Protect(Password);
        }
        public void UnprotectSheet()
        {
            ws.Unprotect();
        }
        public void UnprotectSheet(string Password)
        {
            ws.Unprotect(Password);
        }
        public string[,] ReadRange(int starti, int starty, int endi, int endy)
        {
            Range range = (Range)ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            object[,] holder = range.Value2;
            string[,] returnString = new string[endi - starti + 1, endy - starty + 1];
            for (int p = 1; p <= endi - starti + 1; p++)
            {
                for (int q = 1; q <= endy - starty + 1; q++)
                {
                    if (holder[p, q] != null)
                    {
                        returnString[p - 1, q - 1] = holder[p, q].ToString();
                    }
                    else
                    {
                        returnString[p - 1, q - 1] = " ";
                    }
                }   
            }
            return returnString;
        }
        public void WriteRange(int starti, int starty, int endi, int endy, string[,] writestring)
        {
            Range range = (Range)ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            range.Value2 = writestring;
        }
        public int GetWorksSheetCount(string filePath)
        {
            _Excel.Application excelApp = new _Excel.Application();
            excelApp.Visible = false;

            Workbook workBook = excelApp.Workbooks.Open(filePath);
            int workSheetCount = workBook.Worksheets.Count;

            workBook.Close();
            //workBook.Quit();

            return workSheetCount;
        }
        public static System.Data.DataTable ToDataTable(_Excel.Range range, int rows, int cols)
        {
            System.Data.DataTable table = new System.Data.DataTable();
            for (int i = 1; i <= rows; i++)
            {
                if (i == 1)
                {
                    for (int j = 0; j <= cols; j++)
                    {
                        if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
                            table.Columns.Add(range.Cells[i, j].Value2.ToString());
                        else
                            table.Columns.Add(j.ToString() + ".Sütun");
                    }
                    continue;
                }
                var yeniSatir = table.NewRow();
                for (int j = 1; j <= cols; j++)
                {
                    //Sütunların içeriği boş mu kontrolü yapılmaktadır.
                    if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
                        yeniSatir[j - 1] = range.Cells[i, j].Value2.ToString();
                    else
                        yeniSatir[j - 1] = String.Empty; // İçeriği boş hücrede hata vermesini önlemek için
                }
                table.Rows.Add(yeniSatir);

            }
            return table;
        }
        
    }
}
  