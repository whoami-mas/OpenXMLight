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
        private TableGrid? grid;
        private TableProperties? properties;

        public RowCollection Rows => rows;
        public TableProperties? Properties { get => properties; init
            {
                properties = value;

                this.TableXml.RemoveAllChildren<OpenXML.TableProperties>();
                this.TableXml.AppendChild<OpenXML.TableProperties>(value.TblPropXml);
            } }
        public TableGrid? Grid { get => grid; init 
            {
                grid = value;

                this.TableXml.RemoveAllChildren<OpenXML.TableGrid>();
                this.TableXml.AppendChild<OpenXML.TableGrid>(value.TblGridXml);
            } }


        internal OpenXML.Table TableXml { get; set; }


        public Table()
        {
            TableXml = new OpenXML.Table();

            rows = new RowCollection() { ParentTable = this };

            this.Properties ??= new TableProperties();
            this.Grid ??= new TableGrid();
        }


        internal Table(OpenXML.Table table)
        {
            TableXml = table;

            rows = new RowCollection(table.Elements<OpenXML.TableRow>()) { ParentTable = this };
        }

        public static TableBuilder TableBuilder() => new TableBuilder();
    }
}
