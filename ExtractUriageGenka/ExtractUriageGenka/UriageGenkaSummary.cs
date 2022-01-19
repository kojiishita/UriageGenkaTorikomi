using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtractUriageGenka
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class OrderAttribute : Attribute
    {
        //ソート番号  
        public int Value { get; set; }
    }

    internal class UriageGenkaSummary
    {
        [Order(Value = 1)]
        public int Year { get; set; }
        [Order(Value = 2)]
        public int Month { get; set; }
        [Order(Value = 3)]
        public string Bunrui { get; set; } = null!;
        [Order(Value = 4)]
        public string Busho { get; set; } = null!;
        [Order(Value = 5)]
        public string Kyakusakimei { get; set; } = null!;
        [Order(Value = 6)]
        public string Keiyaku { get; set; } = null!;
        [Order(Value = 7)]
        public string Ankenmei { get; set; } = null!;
        [Order(Value = 8)]
        public decimal Jisseki { get; set; }
        [Order(Value = 9)]
        public decimal JissekZeikomi { get; set; }
    }
}
