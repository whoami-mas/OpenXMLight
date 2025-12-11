using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Vml;
using OpenXMLight.Spreadsheet.Elements;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXML = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Elements.Table
{
    public class Row : ICellObserver
    {
        private CellCollection? _cells = null;

        public CellCollection? Cells
        {
            get => _cells;
            set
            {
                UnsubscribeFromCells();

                //_cells = Skip(value);

                _cells = value;

                SubsribeToCells();

                RowXml?.Append(_cells.Select(s => s.CellXml).ToArray());
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

            _cells = new CellCollection(RowXml.Elements<OpenXML.TableCell>());
            SubsribeToCells();
        }

        internal static HashSet<int> Skip(OpenXML.TableRow rowXml, Cell cell, int mergeOffset)
        {
            var tableCells = rowXml.Elements<OpenXML.TableCell>().ToArray();

            if (mergeOffset < 1)
                throw new ArgumentException("Количество ячеек для объединения должно быть больше 1", nameof(mergeOffset));

            int indexCell = Array.IndexOf(tableCells, cell.CellXml);
            int indexToCell = indexCell + 1 + mergeOffset;

            if (indexToCell > tableCells.Length)
                throw new ArgumentException("Выход за пределы массива");

            HashSet<int> removeCells = new();
            for (int i = indexCell + 1; i < indexToCell; i++)
            {
                cell.Text.Content = cell.Text.Content + tableCells[i].InnerText;

                cell.CellXml.TableCellProperties.TableCellWidth.Width = (int.Parse(cell.CellXml.TableCellProperties.TableCellWidth.Width)
                    + int.Parse(tableCells[i].TableCellProperties.TableCellWidth.Width)).ToString();

                removeCells.Add(i);
                tableCells[i].Remove();
            }

            if (cell.CellXml.TableCellProperties.GridSpan == null)
                cell.CellXml.TableCellProperties.GridSpan = new OpenXML.GridSpan();

            cell.CellXml.TableCellProperties.GridSpan.Val = mergeOffset + 1;

            return removeCells;
        }
        
        internal CellCollection Skip(CellCollection cells)
        {
            var ranges = new List<(int start, int length)>();

            for (int i = 0; i < cells.Count; i++)
            {
                var cell = cells[i];
                if (cell.CellSpan != 0)
                {
                    int start = i + 1;
                    int length = cell.CellSpan - 1;

                    if (start + length <= cells.Count)
                    {
                        ranges.Add((start, length));
                    }
                    else
                        throw new ArgumentException("Выход за границы массива");
                }
            }

            var indicesToExclude = new HashSet<int>();
            foreach (var range in ranges)
            {
                int indexMainCell = range.start - 1;
                for (int i = range.start; i < range.start + range.length; i++)
                {
                    cells[indexMainCell].Text.Content += cells[i].Text.Content;
                    indicesToExclude.Add(i);
                }
            }

            var actualycells = new CellCollection();
            for (int i = 0; i < cells.Count; i++)
            {
                if (!indicesToExclude.Contains(i))
                {
                    actualycells.Add(cells[i]);
                }
            }

            return actualycells;
        }


        #region Subscribe observer
        private void SubsribeToCells()
        {
            if (_cells == null) return;

            foreach (var cell in _cells)
            {
                cell.Row = this;
                cell.AddObserver(this);
            }
        }
        private void UnsubscribeFromCells()
        {
            if (_cells == null) return;

            foreach (var cell in _cells)
            {
                cell.RemoveObserver(this);
            }
        }
        #endregion

        #region observer

        void ICellObserver.OnCellsMerged(HashSet<int> indexRemove)
        {
            for (int i = indexRemove.Max(); i >= indexRemove.Min(); i--)
                _cells.Remove(_cells[i]);
        }

        #endregion
    }
}
