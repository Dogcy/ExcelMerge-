using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergeTools
{
    class BomModel
    {
        public string ParentItemNo { get; set; }
        public string Component { get; set; }
        public List<string> PCBLocation { get; set; }
    }
    class LocationModel
    {
        public int Index { get; set; }
        public string PCBLocationItem { get; set; }
        // 判斷是哪個component寫回
        public string UseComponent { get; set; }
    }
    internal static class LocationC
    {
        public static List<LocationModel> GetLocationModel(string filePath)
        {
            try
            {
                IWorkbook workbook = null;
                FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                if (fs.Name.IndexOf(".xlsx") > 0) // 2007版本 
                {
                    workbook = new XSSFWorkbook(fs); //xlsx數據讀入workbook 
                }
                else if (fs.Name.IndexOf(".xls") > 0) // 2003版本 
                {
                    workbook = new HSSFWorkbook(fs); //xls數據讀入workbook 
                }
                ISheet sheet = workbook.GetSheetAt(0);
                List<LocationModel> locationModels = new List<LocationModel>();
                if (sheet != null)
                {
                    int rowCount = sheet.LastRowNum;
                    // 座標表從0開始計算
                    for (int i = 0; i <= rowCount; i++)
                    {
                        IRow curRow = sheet.GetRow(i);
                        var cellValue0 = curRow.GetCell(0).StringCellValue.Trim();

                        var locationModel = new LocationModel()
                        {
                            Index = i,
                            PCBLocationItem = cellValue0
                        };
                        locationModels.Add(locationModel);
                    }
                }
                return locationModels;
            }
            catch (Exception ex)
            {
                throw new Exception("LocationModel處理有問題 請檢查{座標}檔案內格式", ex);
            }
        }
    }

    internal static class Bom
    {
        public static List<BomModel> GetBomModel(string filePath)
        {

            try
            {


                IWorkbook workbook = null;
                FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                if (fs.Name.IndexOf(".xlsx") > 0) // 2007版本 
                {
                    workbook = new XSSFWorkbook(fs); //xlsx數據讀入workbook 
                }
                else if (fs.Name.IndexOf(".xls") > 0) // 2003版本 
                {
                    workbook = new HSSFWorkbook(fs); //xls數據讀入workbook 
                }
                ISheet sheet = workbook.GetSheetAt(0);
                List<BomModel> boms = new List<BomModel>();
                var checkSeparateChar = string.Empty;
                if (sheet != null)
                {
                    int rowCount = sheet.LastRowNum;
                    for (int i = 1; i <= rowCount; i++)
                    {
                        IRow curRow = sheet.GetRow(i);
                        var cellValue0 = curRow.GetCell(0).StringCellValue.Trim();
                        var cellValue1 = curRow.GetCell(1).StringCellValue.Trim();
                        var cellValue2 = curRow.GetCell(2).StringCellValue.Trim();
                        var bom = new BomModel();
                        bom.ParentItemNo = cellValue0;
                        bom.Component = cellValue1;
                        // 將PCB Location處理
                        if (cellValue2 != "")
                        {
                            if (checkSeparateChar == "")
                            {
                                checkSeparateChar = cellValue2.Any(c => c == ';')
                                    ? ";" : cellValue2.Any(c => c == ':') ?
                                    ":" : cellValue2.Any(c => c == ',') ?
                                    "," : cellValue2.Any(c => c == ' ') ? " " : "";
                            }
                            var Pcbitems = cellValue2.Split(checkSeparateChar).ToList();
                            Pcbitems = Pcbitems.Select(c => c.Replace(" ", "")).ToList();
                            Pcbitems.RemoveAt(Pcbitems.Count - 1); //移除最後一比空字串
                            bom.PCBLocation = Pcbitems;

                            // PCBLocation若為空的不須新增進入
                            boms.Add(bom);

                        }
                    }
                }
                return boms;
            }
            catch (Exception ex)
            {
                throw new Exception("BomModel處理有問題 請檢查{BOM}檔案格式", ex);
            }
        }
    }
    internal static class Merge
    {
        public static bool MergeData(string filePath, List<LocationModel> locations, List<BomModel> bom)
        {
            try
            {

                foreach (var item in locations)
                {
                    var getitem = bom.Where(c => c.PCBLocation.Contains(item.PCBLocationItem)).FirstOrDefault(); //狀況1  有可能會沒有Bom component location資料 所以用null
                    item.UseComponent = getitem?.Component ?? "Null";
                }


                IWorkbook workbook = null;
                FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite);
     
                if (fs.Name.IndexOf(".xlsx") > 0) // 2007版本 
                {
                    workbook = new XSSFWorkbook(fs); //xlsx數據讀入workbook 
                }
                else if (fs.Name.IndexOf(".xls") > 0) // 2003版本 
                {
                    workbook = new HSSFWorkbook(fs); //xls數據讀入workbook 
                }
                ISheet sheet = workbook.GetSheetAt(0);
                foreach (var location in locations)
                {
                    IRow curRow = sheet.GetRow(location.Index);
                    curRow.CreateCell(5).SetCellValue(location.UseComponent);
                }
                FileStream x = File.OpenWrite(filePath);
                workbook.Write(x);//向打開的這個Excel文件中寫入表單並保存。  
                x.Close();

                return true;
            }
            catch (Exception e)
            {
                throw new Exception("資料合併時產生錯誤", e);
            }

        }
    }
}
