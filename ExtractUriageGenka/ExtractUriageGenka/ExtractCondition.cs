using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtractUriageGenka
{
    internal class ExtractCondition
    {
        public string FieldName { get; set; } = null!;
        public string Value { get; set; } = null!;
        public string Operator { get; set; } = null!;
    }
}
