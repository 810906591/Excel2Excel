using Excel2Excel.Common;
using Excel2Excel.Function;
using System;

namespace Excel2Excel.Model
{
    public class DerivedStatistics
    {
        [Excel("流程编号")]
        public string ProcessNumber { get; set; }
        [Excel("创建日期")]
        public DateTime? CreationDate { get; set; }
        [Excel("申请人")]
        public string Claimant { get; set; }
        [Excel("类型")]
        public string Types { get; set; }
        [Excel("衍生料号")]
        public string DerivedPartNumber { get; set; }
        [Excel("产品经理")]
        public string ProductManager { get; set; }
   
    }
    public class DerivedStatisticsRep
    {
        [Excel("流程编号")]
        public string ProcessNumber { get; set; }
        [Excel("创建日期")]
        public DateTime? CreationDate { get; set; }
        [Excel("申请人")]
        public string Claimant { get; set; }
        [Excel("类型")]
        public string Types { get; set; }
        [Excel("衍生料号")]
        public string DerivedPartNumber { get; set; }
        [Excel("产品经理")]
        public string ProductManager { get; set; }
        [Excel("数量")]
        public string Number { get; set; }
        [Excel("金额")]
        public string Amounts { get; set; }

    }
}
