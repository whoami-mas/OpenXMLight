using OpenXMLight.Spreadsheet.Elements;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.Configurations.Elements.Table
{
    public class TableBuilder
    {
        private Table table;

        public TableBuilder SetTableProperties(TableProperties tableProperties)
        {
            table.Properties = tableProperties;
            table.TableXml.AppendChild(tableProperties.TblPropXml);

            return this;
        }
        public TableBuilder SetTableGrid(TableGrid tableGrid)
        {
            table.Grid = tableGrid;
            table.TableXml.AppendChild(tableGrid.TblGridXml);

            return this;
        }

        public TableBuilder AppendRows(params Row[] rows)
        {
            foreach(var row in rows)
                table.Rows.Add(row);

            return this;
        }

        public TableBuilder()
        {
            table = new Table();
        }


        public static implicit operator Table(TableBuilder builder)
        {
            return builder.table;
        }
    }
}
