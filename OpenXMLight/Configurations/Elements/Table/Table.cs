using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXML = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Elements.Table
{
    public class Table
    {
        private RowCollection rows;

        public RowCollection Rows => rows;
        public TableProperties Properties { get; set; }
        public TableGrid Grid { get; set; }


        internal OpenXML.Table TableXml { get; set; }


        public Table() => this.Create();
        public Table(TableProperties? tblProp = default, TableGrid? tblGrid = default) => this.Create(tblProp, tblGrid);


        internal void Create(TableProperties? tblProp = default, TableGrid? tblGrid = default)
        {
            TableXml = new OpenXML.Table();

            rows = new RowCollection() { ParentTable = this };

            //Properties
            Properties = tblProp ?? new TableProperties();
            TableXml.AppendChild(Properties.TblPropXml);

            //Grid
            Grid = tblGrid ?? new TableGrid();
            TableXml.AppendChild(Grid.TblGridXml);
        }



        internal Table(OpenXML.Table table)
        {
            TableXml = table;

            rows = new RowCollection(table.Elements<OpenXML.TableRow>()) { ParentTable = this };
        }
    }
}
