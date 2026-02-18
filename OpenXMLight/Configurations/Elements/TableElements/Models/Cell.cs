using OpenXMLight.Configurations.Elements.Interfaces;
using OpenXMLight.Configurations.Elements.TableElements.Formattings;
using OpenXMLight.Configurations.Elements.TableElements.Formattings.MarginComponents;
using OpenXMLight.Configurations.Elements.TableElements.Formattings.WidthComponents;
using OpenXMLight.Configurations.Formatting;
using OpenXMLight.Tools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Elements.TableElements.Models
{
    public class Cell : Element<OpenXml.TableCell, OpenXml.TableCellProperties>, IObserver
    {
        internal override OpenXml.TableCell ElementXml { get; set; }
        internal override OpenXml.TableCellProperties ElementProperties
        {
            get
            {
                if (_elementProperties == null)
                    _elementProperties = ElementXml.TableCellProperties ??= new OpenXml.TableCellProperties();

                return _elementProperties;
            }
        }


        bool IObserver.IsInitializedCache { get; set; } = false;
        
        internal Cell(OpenXml.TableCell c) => ElementXml = c;



        #region Private properties
        private OpenXml.TableCellProperties? _elementProperties;
        private ElementCollection<Paragraph>? _p;
        private TableCellWidth<CellWidth> _width;
        private TableCellMargin<CellMargin> _margin;
        private TextDirectionType? _textDirection;
        private Color? _color;
        #endregion

        public ElementCollection<Paragraph> Paragraphs
        {
            get
            {
                _p = new(ElementXml.Elements<OpenXml.Paragraph>().Select(s => new Paragraph(s))) { Parent = ElementXml };

                return _p;
            }
        }
        public VerticalAlignments Alignment
        {
            get
            {
                object? value = ElementProperties.TableCellVerticalAlignment?.Val;

                return HelperData.TryParseTableCellVerticalAlignment(value, out VerticalAlignments alignment)
                    ? alignment
                    : Configuration.DEFAULT_VERTICAL_ALIGNMENT;
            }
            set
            {
                ElementProperties.TableCellVerticalAlignment ??= new OpenXml.TableCellVerticalAlignment();
                ElementProperties.TableCellVerticalAlignment.Val = value.Value;
            }
        }
        public int Merged
        {
            get
            {
                if (ElementProperties.GridSpan == null)
                    return 0;

                return ElementProperties.GridSpan.Val;
            }
        }
        public TableCellWidth<CellWidth> Width
        {
            get
            {
                if (_width == null)
                {
                    var width = ElementProperties.TableCellWidth;

                    _width = HelperData.TryParseCellWidth(width, out _width)
                        ? _width
                        : Configuration.DEFAULT_CELLWIDTH;

                    if (width == null)
                        ElementProperties.TableCellWidth = _width._value;
                }

                return _width;
            }
            set
            {
                if (_width == value)
                    return;

                _width = value;

                ElementProperties.TableCellWidth = _width._value;
            }
        }
        public TableCellMargin<CellMargin> Margin
        {
            get
            {
                if(_margin == null)
                {
                    var margin = ElementProperties.TableCellMargin;

                    _margin = new(margin);

                    if (margin == null)
                        ElementProperties.TableCellMargin = _margin._value;
                }

                return _margin;
            }
        }
        public TextDirectionType? TextDirection
        {
            get
            {
                object? _tmpTextDirection = ElementProperties.TextDirection?.Val?.Value;

                _textDirection = HelperData.TryParseTextDirectionCell(_tmpTextDirection, out TextDirectionType? _result)
                    ? _result
                    : Configuration.DEFAULT_TEXTDIRECTION;

                return _textDirection;
            }
            set
            {
                if (TextDirection == value)
                    return;

                _textDirection = value;

                if (_textDirection == null)
                {
                    ElementProperties?.TextDirection?.Remove();
                    return;
                }

                ElementProperties.TextDirection ??= new OpenXml.TextDirection();
                ElementProperties.TextDirection.Val = _textDirection.Value.Value;
            }
        }
        public Color? Color { 
            get
            {
                object? shdColor = ElementProperties.Shading;

                _color = HelperData.TryParseColorShade(shdColor, out Color? _result)
                    ? _result
                    : Configuration.DEFAULT_COLOR_SHADE;

                return _color;
            }
            set
            {
                if (Color == value)
                    return;

                _color = value;

                if (_color == null)
                {
                    ElementProperties.Shading.Remove();
                    return;
                }

                ElementProperties.Shading ??= new();
                ElementProperties.Shading.Val = OpenXml.ShadingPatternValues.Clear;
                ElementProperties.Shading.Color = "auto";
                ElementProperties.Shading.Fill = _color.Value.Hex;
            }
        }


        void IObserver.RefreashCached()
        {
            ((IObserver)this).IsInitializedCache = false;
        }
    }
}
