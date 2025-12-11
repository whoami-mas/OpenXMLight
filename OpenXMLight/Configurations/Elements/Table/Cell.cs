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
        private VerticalMerge vMerge;

        List<ICellObserver> observers = new();

        public Row? Row { get; set; }

        public Text Text
        {
            get => text;
            init
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
                CellXml.TableCellProperties?.AppendChild(new OpenXML.TableCellWidth() {Type = OpenXML.TableWidthUnitValues.Dxa, Width = width.ToString() });
            }
        }
        public int CellSpan => this.CellXml.TableCellProperties?.GridSpan?.Val ?? 0;
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


        public Cell() => this.Create();
        
        public Cell(Text text,int width = 0, VerticalMerge vMerge = VerticalMerge.Non)
        {
            Create();

            this.Text = text;
            this.VMerge = vMerge;
            this.Width = width;
        }

        internal Cell(OpenXML.TableCell cell)
        {
            CellXml = cell;

            this.Text = new Text(cell.Elements<OpenXML.Paragraph>().First());
            this.width = int.Parse(cell.TableCellProperties?.TableCellWidth?.Width);
            
            if (cell.TableCellProperties?.VerticalMerge != null)
            {
                if (cell.TableCellProperties?.VerticalMerge.Val == OpenXML.MergedCellValues.Restart)
                    this.vMerge = VerticalMerge.Start;
                else
                    this.vMerge = VerticalMerge.Continue;
            }
            else
                this.vMerge = VerticalMerge.Non;
        }


        private void Create()
        {
            CellXml = new OpenXML.TableCell();

            CellXml.Append(
                //new OpenXML.Paragraph(),
                new OpenXML.TableCellProperties()
                );
        }

        public Cell Merge(int mergeOffset)
        {
            OpenXML.TableRow? parentRow = this.CellXml.Parent as OpenXML.TableRow;

            HashSet<int> hashIndexRemove = new();
            if (parentRow != null)
            {
                hashIndexRemove = Row.Skip(parentRow, this, mergeOffset);

                NotifyObserver(hashIndexRemove);
            }
            else
            {
                this.CellXml.TableCellProperties.GridSpan ??= new OpenXML.GridSpan();
                this.CellXml.TableCellProperties.GridSpan.Val = mergeOffset + 1;
            }

            return this;
        }

        #region observers
        
        internal void AddObserver(ICellObserver observer)
        {
            if(!observers.Contains(observer))
                observers.Add(observer);
        }
        internal void RemoveObserver(ICellObserver observer)
        {
            observers.Remove(observer);
        }
        internal void NotifyObserver(HashSet<int> indexCellRemove)
        {
            foreach (var observer in observers)
                observer.OnCellsMerged(indexCellRemove);
        }
        #endregion
    }
}
