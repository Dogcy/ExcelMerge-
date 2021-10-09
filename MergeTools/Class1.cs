using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergeTools
{
    internal class Class1
    {
        /// <summary>
        /// epplus套件無法使用.xls檔所以放置這做紀錄
        /// </summary>
        public void EPPlus() {
            var excelFile = new FileInfo("file");


            using (ExcelPackage package = new ExcelPackage(excelFile))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                var sheet = package.Workbook.Worksheets[0];

                int colCount = sheet.Dimension.End.Column;  //get Column Count
                int rowCount = sheet.Dimension.End.Row;     //get row count
                for (int row = 1; row <= rowCount; row++)
                {
                    for (int col = 1; col <= colCount; col++)
                    {
                        Console.WriteLine(" Row:" + row + " column:" + col + " Value:" + sheet.Cells[row, col].Value?.ToString().Trim());
                    }
                }

                //var boms = GetList<BOM>(sheet);
            }
        }
    }
}
