using System.Collections.Generic;

namespace Excel2Excel.Model
{
    /// <summary>
    /// 文件转实体类条件类
    /// </summary>
    public class FileMapUtil
    {
        /// <summary>
        /// 无参构造函数
        /// </summary>
        public FileMapUtil() { }
        /// <summary>
        /// 根据列头和对应实体类字段名创建帮助类，列数默认为0
        /// </summary>
        /// <param name="columnHeader">列头</param>
        /// <param name="fieldName">字段名</param>
        public FileMapUtil(string columnHeader, string fieldName)
        {
            ColumnHeader = columnHeader;
            FieldName = fieldName;
        }
        /// <summary>
        /// 根据列头和对应实体类字段名创建帮助类，列数默认为0
        /// </summary>
        /// <param name="columnHeader">列头</param>
        /// <param name="fieldName">字段名</param>
        public FileMapUtil(string columnHeader, string fieldName, Dictionary<string, int> map)
        {
            ColumnHeader = columnHeader;
            FieldName = fieldName;
            Map = map;
        }
        /// <summary>
        /// 文件列头
        /// </summary>
        public string ColumnHeader { get; set; }
        /// <summary>
        /// 此列对应列数
        /// </summary>
        public int LineNumber { get; set; }
        /// <summary>
        /// 对应实体类字段名
        /// </summary>
        public string FieldName { get; set; }

        /// <summary>
        /// 状态集合
        /// </summary>
        public Dictionary<string, int> Map { get; set; }
    }
}
