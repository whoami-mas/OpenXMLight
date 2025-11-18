using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXMLight.Configurations.Formatting;
using OpenXML = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Elements.Table
{
    public class Cell
    {
        private Text text;
        private int width;
        private int mergeColumn;
        private VerticalMerge vMerge;

        public Text Text
        {
            get => text;
            set
            {
                text = value;

                CellXml.RemoveAllChildren<OpenXML.Paragraph>();
                CellXml.AppendChild(text.Properties.Paragraph);
            }
        }
        public int Width
        {
            get => width;
            set
            {
                width = value;

                CellXml.TableCellProperties?.RemoveAllChildren<OpenXML.TableCellWidth>();
                CellXml.TableCellProperties?.AppendChild(new OpenXML.TableCellWidth() {Type = OpenXML.TableWidthUnitValues.Pct, Width = width.ToString() });
            }
        }
        public int MergeColumn
        {
            get => mergeColumn;
            set
            {
                mergeColumn = value;

                if(mergeColumn >= 1)
                {
                    CellXml.TableCellProperties?.RemoveAllChildren<OpenXML.GridSpan>();
                    CellXml.TableCellProperties?.AppendChild(new OpenXML.GridSpan() { Val = mergeColumn });
                }
            }
        }
        public VerticalMerge VMerge
        {
            get => vMerge;
            set
            {
                vMerge = value;

                switch(vMerge)
                {
                    case VerticalMerge.Start:
                        CellXml.TableCellProperties?.RemoveAllChildren<OpenXML.VerticalMerge>();
                        CellXml.TableCellProperties?.AppendChild(new OpenXML.VerticalMerge() { Val = OpenXML.MergedCellValues.Restart });
                        break;
                    case VerticalMerge.Continue:
                        CellXml.TableCellProperties?.RemoveAllChildren<OpenXML.VerticalMerge>();
                        CellXml.TableCellProperties?.AppendChild(new OpenXML.VerticalMerge());
                        break;
                    case VerticalMerge.Non:
                        break;
                }
            }
        }

        internal OpenXML.TableCell CellXml { get; set; }



        public Cell() => this.CreateCell();
        
        public Cell(Text text, int mergeColumn = 0, VerticalMerge vMerge = VerticalMerge.Non)
        {
            CreateCell();

            this.Text = text;
            this.MergeColumn = mergeColumn;
            this.VMerge = vMerge;
        }

        internal Cell(OpenXML.TableCell cell)
        {
            CellXml = cell;
        }


        private void CreateCell()
        {
            CellXml = new OpenXML.TableCell();

            CellXml.Append(
                new OpenXML.Paragraph(),
                new OpenXML.TableCellProperties()
                );
        }
    }
}
