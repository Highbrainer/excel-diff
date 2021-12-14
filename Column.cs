using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDiff
{
    class Column
    {
        public string Name { get; set; }
        public int index { get; set; }

        public Column(int index)
        {
            this.index = index;
            this.Name = ""+(char)(((int)'A') + index);
        }
        public override string ToString()
        {
            return Name;
        }
    }
}
