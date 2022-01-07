using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtractUriageGenka
{
    internal class UriageGenkaSummary
    {
        public int Year { get; set; }
        public int Month { get; set; }
        public string Bunrui { get; set; } = null!;
        public string Busho { get; set; } = null!;
        public string Kyakusakimei { get; set; } = null!;
        public string Keiyaku { get; set; } = null!;
        public string Ankenmei { get; set; } = null!;
        public decimal Jisseki { get; set; }
        public decimal JissekZeikomi { get; set; }
    }
}
