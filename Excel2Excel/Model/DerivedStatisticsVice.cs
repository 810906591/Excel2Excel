using Excel2Excel.Function;

namespace Excel2Excel.Model
{
    class DerivedStatisticsVice
    {
        /// <summary>
        /// 标记型号
        /// </summary>
        [Excel("型号")]
        public string TagType { get; set; }
        [Excel("物料编码")]
        public string DerivedPartNumber { get; set; }
        [Excel("物料名称")]
        public string MaterialName { get; set; }
        [Excel("规格型号")]
        public string Specifications { get; set; }

        //public string Types { get; set; }

        //public string Units { get; set; }
        [Excel("数量")]
        public string Number { get; set; }
        [Excel("金额")]
        public string Amounts { get; set; }

    }
}
