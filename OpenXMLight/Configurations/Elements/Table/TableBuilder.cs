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
        public TableBuilder SetTableGrid(params int[] widthColumn)
        {
            int countCell = (int)table.Rows?.Select(s => s?.Cells?.Count).DefaultIfEmpty(0).Max();

            if (countCell != widthColumn.Length)
                throw new ArgumentException("Количество ячеек не сходиться с указанным количеством размера");

            table.Grid.ColumnWidth = widthColumn;
            
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
