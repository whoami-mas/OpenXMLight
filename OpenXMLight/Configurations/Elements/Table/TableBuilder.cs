using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.Configurations.Elements.Table
{
    public class TableBuilder
    {
        private Table Table { get; set; }

        public TableBuilder SetTableGrid(TableGrid tableGRid)
        {
            Table.Grid = tableGRid;

            return this;
        }

        public TableBuilder AppendRows(params Row[] rows)
        {
            foreach(var row in rows)
                Table.Rows.Add(row);

            return this;
        }

        public TableBuilder()
        {
            Table = new();
        }


        public static implicit operator Table(TableBuilder builder)
        {
            return builder.Table;
        }
    }
}
