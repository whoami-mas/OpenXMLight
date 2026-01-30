using OpenXMLight.Configurations.Elements.TableElements.Formattings;
using OpenXMLight.Configurations.Elements.TableElements.Formattings.WidthComponents;
using OpenXMLight.Configurations.Elements.TableElements.Models;
using OpenXMLight.Configurations.Formatting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Elements.TableElements
{
    public class CellBuilder
    {
        private readonly Cell _cell;



        public CellBuilder() : this(new OpenXml.TableCell())
        {

        }
        internal CellBuilder(OpenXml.TableCell c)
        {
            _cell = new Cell(c);
        }



        public CellBuilder AddParagraph(Action<ParagraphBuilder>? configuration = null)
        {
            OpenXml.Paragraph p = new OpenXml.Paragraph();
            _cell.ElementXml.AppendChild(p);

            var paragraph = new ParagraphBuilder(p);
            configuration?.Invoke(paragraph);

            return this;
        }
        public CellBuilder SetVerticalAlignment(VerticalAlignments alignment)
        {
            _cell.Alignment = alignment;

            return this;
        }
        public CellBuilder SetWidth(Action<TableCellWidth<CellWidth>>? configuration = null)
        {
            TableCellWidth<CellWidth> width = new();
            configuration?.Invoke(width);

            _cell.Width = width;

            return this;
        }
        public CellBuilder SetTextDirection(TextDirectionType textDirection)
        {
            _cell.TextDirection = textDirection;

            return this;
        }



        public static implicit operator Cell(CellBuilder build) => build._cell;
    }
}
