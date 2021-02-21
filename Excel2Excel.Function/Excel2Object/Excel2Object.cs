using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
//https://www.cnblogs.com/hydor/p/3812248.html
namespace Excel2Excel.Function
{
    public class Excel2Object
    {
        public class ExcelHelper
        {
            public static IEnumerable<TModel> ExcelToObject<TModel>(string path, string sheetname) where TModel : class, new()
            {
                var importer = new ExcelImporter();
                return importer.ExcelToObject<TModel>(path, sheetname);
            }

            public static void ObjectToExcel<TModel>(IEnumerable<TModel> data, string path) where TModel : class, new()
            {
                var importer = new ExcelExporter();
                var bytes = importer.ObjectToExcelBytes(data);
                File.WriteAllBytes(path, bytes);
            }
        }
        internal class ExcelImporter
        {
            public IEnumerable<TModel> ExcelToObject<TModel>(string path, string sheetname) where TModel : class, new()
            {
                var result = GetDataRows(path, sheetname);
                var dict = ExcelUtil.GetExportAttrDirt<TModel>();
                var dictColums = new Dictionary<int, KeyValuePair<PropertyInfo, ExcelAttribute>>();
                IEnumerator rows = result;
            dc: var titleRow = (IRow)rows.Current;
                if (titleRow != null)
                {
                    foreach (var cell in titleRow.Cells)
                    {
                        var prop = new KeyValuePair<PropertyInfo, ExcelAttribute>();
                        foreach (var item in dict)
                        {
                            if (cell.StringCellValue == item.Value.Title)
                                prop = item;
                        }
                        if (prop.Key != null && !dictColums.ContainsKey(cell.ColumnIndex))
                            dictColums.Add(cell.ColumnIndex, prop);
                    }
                    //遍历
                    if (dictColums.Count <= 0)
                    {
                        rows.MoveNext();
                        goto dc;
                    }
                }
                while (rows.MoveNext())
                {
                    var row = (IRow)rows.Current;
                    if (row != null)
                    {
                        var firstCell = row.GetCell(0);
                        if (firstCell == null || firstCell.CellType == CellType.Blank
                            || string.IsNullOrEmpty(firstCell.ToString()))
                            continue;
                    }
                    var model = new TModel();
                    foreach (var pair in dictColums)
                    {
                        var propType = pair.Value.Key.PropertyType;
                        if (propType == typeof(DateTime?) || propType == typeof(DateTime))
                            pair.Value.Key.SetValue(model, GetCellDateTime(row, pair.Key), null);
                        else
                        {
                            try
                            {
                                var val = Convert.ChangeType(GetCellValue(row, pair.Key), propType);
                                pair.Value.Key.SetValue(model, val, null);
                            }
                            catch (Exception) { }
                        }
                    }
                    yield return model;
                }
            }
            /// <summary>
            /// 获取值
            /// </summary>
            /// <param name="row"></param>
            /// <param name="index"></param>
            /// <returns></returns>
            string GetCellValue(IRow row, int index)
            {
                var result = string.Empty;
                try
                {
                    switch (row.GetCell(index).CellType)
                    {
                        case CellType.Numeric:
                            result = row.GetCell(index).NumericCellValue.ToString();
                            break;
                        case CellType.String:
                            result = row.GetCell(index).StringCellValue.ToString();
                            break;
                        case CellType.Blank:
                            break;
                        case CellType.Formula:
                            result = row.GetCell(index).CellFormula;
                            break;
                        case CellType.Boolean:
                            result = row.GetCell(index).NumericCellValue.ToString();
                            break;
                        case CellType.Unknown:
                            result = row.GetCell(index).NumericCellValue.ToString();
                            break;
                        default:
                            result = row.GetCell(index).ToString();
                            break;
                    }
                }
                catch (Exception) { }
                return result;
            }
            /// <summary>
            /// 获取某工作簿数据
            /// </summary>
            /// <param name="path"></param>
            /// <param name="sheetname"></param>
            /// <returns></returns>
            IEnumerator GetDataRows(string path, string sheetname)
            {
                if (string.IsNullOrEmpty(path)) return null;
                IWorkbook workbook = null;
                try
                {
                    using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read))
                    {
                        if (path.IndexOf(".xlsx") > 0) // 2007版本
                            workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook

                        else if (path.IndexOf(".xls") > 0) // 2003版本
                            workbook = new HSSFWorkbook(fileStream);  //xls数据读入workbook
                    }
                    ISheet sheet = workbook.GetSheet(sheetname);
                    IEnumerator rows = sheet.GetRowEnumerator();
                    rows.MoveNext();
                    return rows;
                }
                catch (Exception) { return null; }
            }

            /// <summary>
            /// 获取时间
            /// </summary>
            /// <param name="row"></param>
            /// <param name="index"></param>
            /// <returns></returns>
            #region MyRegion
            DateTime? GetCellDateTime(IRow row, int index)
            {
                DateTime? result = null;
                try
                {
                    switch (row.GetCell(index).CellType)
                    {
                        case CellType.Numeric:
                            try
                            {
                                result = row.GetCell(index).DateCellValue;
                            }
                            catch (Exception) { }
                            break;
                        case CellType.String:
                            var str = row.GetCell(index).StringCellValue;
                            if (str.EndsWith("年"))
                            {
                                if (DateTime.TryParse((str + "-01-01").Replace("年", ""), out DateTime dt))
                                    result = dt;
                            }
                            else if (str.EndsWith("月"))
                            {
                                if (DateTime.TryParse((str + "-01").Replace("年", "").Replace("月", ""), out DateTime dt))
                                    result = dt;
                            }
                            else if (!str.Contains("年") && !str.Contains("月") && !str.Contains("日"))
                            {
                                try
                                {
                                    result = Convert.ToDateTime(str);
                                }
                                catch (Exception)
                                {
                                    try
                                    {
                                        result = Convert.ToDateTime((str + "-01-01").Replace("年", "").Replace("月", ""));
                                    }
                                    catch (Exception) { result = null; }
                                }
                            }
                            else
                            {
                                if (DateTime.TryParse(str.Replace("年", "").Replace("月", ""), out DateTime dt))
                                    result = dt;
                            }
                            break;
                        case CellType.Blank:
                            break;
                    }
                }
                catch (Exception)
                {

                }
                return result;
            }
            #endregion

        }
        internal class ExcelExporter
        {
            public byte[] ObjectToExcelBytes<TModel>(IEnumerable<TModel> data)
            {
                var workbook = new HSSFWorkbook();
                var sheet = workbook.CreateSheet();
                var attrDict = ExcelUtil.GetExportAttrDirt<TModel>();
                var attrArray = new KeyValuePair<PropertyInfo, ExcelAttribute>[attrDict.Count];
                int aNum = 0;
                foreach (var item in attrDict)
                {
                    attrArray[aNum] = item;
                    aNum++;
                }
                for (int i = 0; i < attrArray.Length; i++)
                {
                    sheet.SetColumnWidth(i, 50 * 50);
                }
                var headerRow = sheet.CreateRow(0);
                for (int i = 0; i < attrArray.Length; i++)
                {
                    headerRow.CreateCell(i).SetCellValue(attrArray[i].Value.Title);
                }
                int rowNumber = 1;
                foreach (var item in data)
                {
                    var row = sheet.CreateRow(rowNumber++);
                    
                    for (int i = 0; i < attrArray.Length; i++)
                    {
                        row.CreateCell(i).SetCellValue((attrArray[i].Key.GetValue(item, null) ?? "").ToString());
                    }
                }
                using (var output = new MemoryStream())
                {
                    workbook.Write(output);
                    var bytes = output.ToArray();
                    return bytes;
                }
            }
           
        }
        internal class ExcelUtil
        {
            public static Dictionary<PropertyInfo, ExcelAttribute> GetExportAttrDirt<T>()
            {
                var dict = new Dictionary<PropertyInfo, ExcelAttribute>();
                foreach (var p in typeof(T).GetProperties())
                {
                    var attr = new object();
                    var ppi = p.GetCustomAttributes(true);
                    for (int i = 0; i < ppi.Length; i++)
                    {
                        if (ppi[i] is ExcelAttribute)
                        {
                            attr = ppi[i];
                            break;
                        }
                    }
                    if (attr != null)
                        dict.Add(p, attr as ExcelAttribute);
                }
                return dict;
            }

        }
    }
}
