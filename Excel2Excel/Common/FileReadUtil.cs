namespace Excel2Excel.Common
{
    public class FileReadUtil
    {

        /// <summary>
        /// 公用文件读取类
        /// </summary>
        /// <typeparam name="T">泛型，具体要转换成的类</typeparam>
        /// <param name="pojoList">字段名和表列头映射关系集合</param>
        /// <param name="fieldName">一个绝对不为空的字段，用来判断当前记录是否有效，部分文件会出现空行问题！</param>
        /// <param name="filePath">文件绝对路径（跟项目在同一服务器上）</param>
        /// <returns></returns>
        //public static List<T> GM<T>(List<FileMapUtil> pojoList, string sheetName, string fieldName, string filePath)
        //{
        //    //filePath = "F:\\demo.xlsx";用作本地测试
        //    //创建泛型集合
        //    List<T> tList = new List<T>();
        //    IWorkbook workbook = null;  //新建IWorkbook对象
        //    FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
        //    if (filePath.IndexOf(".xlsx") > 0) // 2007版本
        //    {
        //        workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook
        //    }
        //    else if (filePath.IndexOf(".xls") > 0) // 2003版本
        //    {
        //        workbook = new HSSFWorkbook(fileStream);  //xls数据读入workbook
        //    }

        //    //ISheet sheet = workbook.GetSheetAt(0);  //获取第一个工作表
        //    ISheet sheet = workbook.GetSheet(sheetName);
        //    IRow row;// = sheet.GetRow(0);            //新建当前工作表行数据
        //    bool result = false;                    //用于中断循环
        //    for (int i = 0; i <= sheet.LastRowNum; i++)
        //    {
        //        if (result)//是否中断循环
        //            break;

        //        //创建泛型对象
        //        T item = System.Activator.CreateInstance<T>();
        //        Type type = item.GetType();
        //        //获取所有公共字段名
        //        PropertyInfo[] propertyInfoList = type.GetProperties();
        //        row = sheet.GetRow(i);   //row读入第i行数据
        //        //Console.WriteLine(row);
        //        if (row != null)
        //        {
        //            //对工作表每一列进行读取
        //            for (int j = 0; j < row.LastCellNum; j++)
        //            {
        //                //i = 0 表明读取列头
        //                if (i == 0)
        //                {
        //                    //循环遍历传进来的实体类文件映射类
        //                    for (int k = 0; k < pojoList.Count; k++)
        //                    {
        //                        //判断是否读取到了内容
        //                        if (row.GetCell(j) == null)
        //                        {
        //                            continue;
        //                        }
        //                        string rowName = rowName = row.GetCell(j).ToString();
        //                        //判断列名是否和映射类中的列名相同
        //                        if (pojoList[k].ColumnHeader.Equals(rowName))
        //                        {
        //                            pojoList[k].LineNumber = j;
        //                        }
        //                    }
        //                    continue;
        //                }
        //                //读取表内容,循环遍历传进来的实体类文件映射类
        //                for (int k = 0; k < pojoList.Count; k++)
        //                {
        //                    if (result)
        //                        break;
        //                    //判断是否读取到了内容
        //                    if (row.GetCell(j) == null || row.GetCell(j).ToString() == "")
        //                        continue;

        //                    string rowValue = row.GetCell(j).ToString();
        //                    //判断列数是否和文件列数相同，如果相同值就赋给对象
        //                    if (pojoList[k].LineNumber == j)
        //                    {
        //                        //遍历泛型类中的所有字段
        //                        foreach (PropertyInfo prop in propertyInfoList)
        //                        {
        //                            //Console.WriteLine(prop.Name);
        //                            //如果字段和我们定义的实体类字段名相同则进行赋值操作
        //                            if (prop.Name.Equals(pojoList[k].FieldName))
        //                            {
        //                                //Console.WriteLine(prop.PropertyType.Name);
        //                                //如果这一列为状态列，则进行状态码的获取填充
        //                                if (pojoList[k].Map != null)
        //                                {
        //                                    Dictionary<string, int> dic = pojoList[k].Map;
        //                                    //如果这个文件中的状态列是我们定义的其中一种则进行封装，否则设置值为3用于后期的
        //                                    if (dic.ContainsKey(rowValue))
        //                                    {
        //                                        prop.SetValue(item, Convert.ToInt32(dic[rowValue]), null);
        //                                    }
        //                                    else
        //                                    {
        //                                        prop.SetValue(item, 3, null);
        //                                        result = true;//中断操作
        //                                    }
        //                                    break;
        //                                }
        //                                //进行一个类型的判断，如果你还有其他类型，请自行添加，不过请先放出上面注释，以保证类型名称不会出错
        //                                //Console.WriteLine(rowValue);
        //                                ConvertByPropTypeName(prop, item, rowValue);
        //                                break; //如果判断是一样的  那么其他的就不用比了
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //        //i == 0 表示当前读取的是表头，不做集合添加操作
        //        if (i != 0)
        //        {
        //            foreach (PropertyInfo prop in propertyInfoList)
        //            {
        //                //Console.WriteLine(fieldName + "   值为：" + prop.GetValue(item, null));
        //                //判断哪个百分百不为0的字段，是否有值，这样就能排除excel的空行
        //                if (prop.Name.Equals(fieldName) && prop.GetValue(item, null) != null && prop.GetValue(item, null).ToString() != "")
        //                {
        //                    tList.Add(item);
        //                }
        //            }
        //        }
        //    }
        //    return tList;
        //}

        ///// <summary>
        ///// 根据字段类型名称进行不同的值转换
        ///// </summary>
        ///// <typeparam name="T">泛型类</typeparam>
        ///// <param name="prop">字段信息</param>
        ///// <param name="item">实体</param>
        ///// <param name="rowValue">文件中的表值</param>
        //private static void ConvertByPropTypeName<T>(PropertyInfo prop, T item, string rowValue)
        //{
        //    if (prop.PropertyType.Name.Equals("DateTime"))
        //    {
        //        prop.SetValue(item, Convert.ToDateTime(rowValue),null);
        //    }
        //    else if (prop.PropertyType.Name.Equals("Int32"))
        //    {
        //        prop.SetValue(item, Convert.ToInt32(rowValue), null);
        //    }
        //    else if (prop.PropertyType.Name.Equals("Decimal"))
        //    {
        //        prop.SetValue(item, Convert.ToDecimal(rowValue), null);
        //    }
        //    else if (prop.PropertyType.Name.Equals("Double"))
        //    {
        //        prop.SetValue(item, Convert.ToDouble(rowValue), null);
        //    }
        //    else
        //    {
        //        prop.SetValue(item, rowValue, null); //设置值
        //    }
        //}

    }
}
