using OpenXMLight.Configurations.Elements.Interfaces;
using OpenXMLight.Configurations.Elements.TableElements.Formattings;
using OpenXMLight.Configurations.Elements.TableElements.Formattings.MarginComponents;
using OpenXMLight.Configurations.Elements.TableElements.Formattings.WidthComponents;
using OpenXMLight.Configurations.Formatting;
using OpenXMLight.Tools;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Elements.TableElements.Models
{
    public class Table : Element<OpenXml.Table, OpenXml.TableProperties>, IObservable
    {
        internal override OpenXml.Table ElementXml { get; set; }
        internal override OpenXml.TableProperties ElementProperties
        {
            get
            {
                if (_elementProperties == null)
                {
                    _elementProperties = ElementXml.Elements<OpenXml.TableProperties>().FirstOrDefault();

                    if (_elementProperties == null)
                        _elementProperties = ElementXml.PrependChild(new OpenXml.TableProperties());
                }


                return _elementProperties;
            }
        }


        List<IObserver> observers;

        internal Table(OpenXml.Table tbl) => ElementXml = tbl;


        #region Private properties
        private OpenXml.TableProperties? _elementProperties;
        private ElementCollection<Row>? _rows;
        private BordersLine? _borders;
        private TableCellWidth<TableWidth> _width;
        private TableCellMargin<TableMargin> _margin;
        #endregion

        public ElementCollection<Row> Rows
        {
            get
            {
                if (_rows == null)
                {
                    observers = new();

                    _rows = new(ElementXml.Elements<OpenXml.TableRow>().Select(s => new Row(s))) { Parent = ElementXml };

                    observers.AddRange(_rows);
                }

                return _rows;
            }
        }
        public BordersLine? Borders
        {
            get
            {
                if (_borders == null)
                {
                    OpenXml.TableBorders? tblBorders = ElementProperties.TableBorders;

                    _borders = HelperData.TryParseTableBorders(tblBorders, out _borders)
                        ? _borders
                        : Configuration.DEFAULT_BORDERS_LINE;

                }

                return _borders;
            }
            set
            {
                if (_borders == value)
                    return;

                _borders = value;

                ElementProperties.TableBorders = value;
            }
        }
        public TableCellWidth<TableWidth> Width
        {
            get
            {
                if (_width == null)
                {
                    var width = ElementProperties.TableWidth;

                    _width = HelperData.TryParseTableWidth(width, out _width)
                        ? _width
                        : Configuration.DEFAULT_TABLEWIDTH;

                    if (width == null)
                        ElementProperties.AppendChild<OpenXml.TableWidth>(_width._value);
                }

                return _width;
            }
            set
            {
                if (_width == value)
                    return;

                _width = value;

                ElementProperties.TableWidth = _width._value;
            }
        }
        public bool IsFixed
        {
            get
            {
                OpenXml.TableLayout? layout = ElementProperties.TableLayout;

                if (layout == null)
                    return false;

                if (layout.Type == OpenXml.TableLayoutValues.Fixed)
                    return true;
                else
                    return false;
            }
            set
            {
                ElementProperties.TableLayout ??= new OpenXml.TableLayout();

                if (value)
                    ElementProperties.TableLayout.Type = OpenXml.TableLayoutValues.Fixed;
                else
                    ElementProperties.TableLayout.Remove();
            }
        }
        public int CountColumn => Rows.Select(s => s.Cells.Count).DefaultIfEmpty(0).Max();
        public TableCellMargin<TableMargin> Margin 
        {
            get
            {
                if(_margin == null)
                {
                    var margin = ElementProperties.TableCellMarginDefault;

                    _margin = new(margin);

                    if (margin == null)
                        ElementProperties.TableCellMarginDefault = _margin._value;
                }

                return _margin;
            }
            protected set
            {

            } 
        }


        public void MergeCell(int c1_r, int c1_c,
            int c2_r, int c2_c)
        {
            if (c1_r == c2_r && c1_c == c2_c)
                return;

            if (c1_r < 0 || c1_c < 0 || c2_r < 0 || c2_c < 0)
                throw new ArgumentException("Выход за пределы границы");

            if (c1_r > c2_r || c1_c > c2_c)
                throw new ArgumentException("Начальная ячейка не может иметь индексы больше конечной");

            if (c2_r > Rows.Count)
                throw new ArgumentException("Количество строк меньше чем конечная ячейка");
            if (c2_c > CountColumn)
                throw new ArgumentException("Количество колонок меньше чем конечная ячейка");

            int indexFirstRow = c1_r;
            for (int i = c1_r; i <= c2_r; i++)
            {
                Row row = Rows[i];
                int indexPost = c2_c - 1;
                int valSpan = 1;
                for (int j = c2_c; j != c1_c; j--)
                {
                    foreach (Paragraph paragraph in row.Cells[j].Paragraphs)
                        if(!string.IsNullOrWhiteSpace(paragraph.AllText))
                            row.Cells[indexPost].Paragraphs.Add(paragraph, true);

                    row.Cells[j].ElementXml.Remove();
                    indexPost--;
                    valSpan++;
                }

                indexPost = indexPost + 1;
                row.Cells[indexPost].ElementProperties.GridSpan ??= new OpenXml.GridSpan();
                row.Cells[indexPost].ElementProperties.GridSpan.Val = valSpan;

                if (c1_r == c2_r)
                    break;

                if (i == c1_r)
                {
                    row.Cells[indexPost].ElementProperties.VerticalMerge ??= new OpenXml.VerticalMerge();
                    row.Cells[indexPost].ElementProperties.VerticalMerge.Val = OpenXml.MergedCellValues.Restart;
                }
                else
                {
                    foreach (Paragraph paragraph in row.Cells[indexPost].Paragraphs)
                        if (!string.IsNullOrWhiteSpace(paragraph.AllText))
                            Rows[indexFirstRow].Cells[indexPost].Paragraphs.Add(paragraph, true);

                    row.Cells[indexPost].Paragraphs.Clear();

                    row.Cells[indexPost].ElementProperties.VerticalMerge ??= new OpenXml.VerticalMerge();
                }
            }

            ((IObservable)this).NotifyObservers();
        }

        void IObservable.NotifyObservers()
        {
            foreach (var observer in observers)
                observer.RefreashCached();
        }
    }
}
