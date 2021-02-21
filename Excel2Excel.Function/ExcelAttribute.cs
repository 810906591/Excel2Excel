using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Excel2Excel.Function
{
    public class ExcelAttribute : Attribute
    {
        public ExcelAttribute(string name)
        {
            Title = name;
        }
        public string Title { get; set; }
    }
}
