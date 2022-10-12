using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadWriteExcel
{
    class Program
    {
        static void Main(string[] args)
        {

            Console.OutputEncoding = Encoding.Unicode;
            Console.InputEncoding = Encoding.Unicode;

            ReadExcel();
            //WriteExcel();
            Console.Read();

        }

        public static void ReadExcel()
        {
            
            Excel.Application excel = new Excel.Application();
            Excel.Application excel2 = new Excel.Application();

            Excel.Workbook excelWorkbook2 = excel2.Workbooks.Add();
            Excel.Worksheet excelWorksheet2 = (Excel.Worksheet)excelWorkbook2.Sheets.Add();

            Excel.Workbook xlWorkbook = excel.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory +  "Import_Cur.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range excelRange = xlWorksheet.UsedRange;

            int rowCount = excelRange.Rows.Count;
            int colCount = excelRange.Columns.Count;
            int row2 = 2;

            for (int i = 1; i < colCount-1;i++)
            {
                excelWorksheet2.Cells[1, i] = excelRange.Cells[6, i].Value2.ToString();
            }

            String tacGiaTruong = excelRange.Cells[7, 11].Value2.ToString();
            //Console.Write(tacGiaTruonn);

            for (int i = 7; i <= rowCount; i++)
            {
                String tmp = null;
                if (excelRange.Cells[i, 3] != null && excelRange.Cells[i, 3].value2 != null) 
                    tmp = excelRange.Cells[i, 3].Value2.ToString();
                if (String.Equals(tmp, tacGiaTruong))
                {
                    for (int j = 1; j < colCount; j++)
                    {
                        if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].value2 != null)
                            excelWorksheet2.Cells[row2, j] = excelRange.Cells[i, j].Value2.ToString();
                    }
                    row2++;
                 }
            }

            excel2.ActiveWorkbook.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "Export_" + DateTime.Now.ToString("dd_MMMM_hh_mm_ss_tt") + ".xlsx", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlShared, false, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);

            excelWorkbook2.Close();
            excel2.Quit();

            Marshal.FinalReleaseComObject(excelWorksheet2);
            Marshal.FinalReleaseComObject(excelWorkbook2);
            Marshal.FinalReleaseComObject(excel2);

            GC.Collect();
            GC.WaitForPendingFinalizers();
            Console.WriteLine("Đã ghi file xong!");


            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(excelRange);
            Marshal.ReleaseComObject(xlWorksheet);
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            excel.Quit();
            Marshal.ReleaseComObject(excel);
            Console.WriteLine("Đã đọc file xong!");
        }

    }
}
