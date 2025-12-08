using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXML = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Elements.Table
{
    public class TableGrid
    {
        private int[]? columnWidth;
        

        public int[]? ColumnWidth 
        {
            get => columnWidth;
            set
            {
                columnWidth = value;

                TblGridXml.RemoveAllChildren<OpenXML.GridColumn>();
                TblGridXml.Append(
                    columnWidth
                            .Select((value, index) => new OpenXML.GridColumn()
                            {
                                Width = (value * Configuration.TwipsInPixels).ToString()
                            })
                            .ToArray()
                );
            }
        }


        internal OpenXML.TableGrid TblGridXml { get; set; }
        
        
        public TableGrid() => this.Create();


        internal void Create()
        {
            TblGridXml = new OpenXML.TableGrid();
        }
    }
}
