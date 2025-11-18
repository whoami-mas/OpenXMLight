using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXML = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Elements.Table
{
    public class Row
    {
        private CellCollection _cells = null;

        public CellCollection Cells
        {
            get => _cells;
            set
            {
                _cells = value;

                RowXml?.Append(value.Select(s => s.CellXml).ToArray());
            }
        }
        
        internal OpenXML.TableRow RowXml { get; set; } = new OpenXML.TableRow();
        
        public Row()
        {
            //Cells = new CellCollection();
        }

        internal Row(OpenXML.TableRow row)
        {
            RowXml = row;

            Cells = new CellCollection();
        }
    }
}
