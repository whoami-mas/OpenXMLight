using OpenXMLight.Configurations.Elements.Interfaces;
using OpenXMLight.Configurations.Elements.TableElements.Formattings.FormattingsBase;
using OpenXMLight.Configurations.Elements.TableElements.Formattings.MarginComponents.Margins;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml;
using OpenXmlElement = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Elements.TableElements.Formattings.MarginComponents
{
    public class TableMargin : IMarginContainer,
        IElementBase<OpenXmlElement.TableCellMarginDefault>
    {
        public OpenXmlElement.TableCellMarginDefault ElementXml { get; set; }
        public Margin<TopMargin> Top { get; protected set; }
        public Margin<BottomMargin> Bottom { get; protected set; }
        public Margin<TableCellLeftMargin> Left { get; protected set; }
        public Margin<TableCellRightMargin> Right { get; protected set; }


        IMargin IMarginContainer.Top => Top;
        IMargin IMarginContainer.Bottom => Bottom;
        IMargin IMarginContainer.Left => Left;
        IMargin IMarginContainer.Right => Right;

        public void Initialize(OpenXml.OpenXmlElement? xml = null)
        {
            ElementXml = xml as OpenXmlElement.TableCellMarginDefault ?? new();

            Top = new(ElementXml.TopMargin);
            Bottom = new(ElementXml.BottomMargin);
            Left = new(ElementXml.TableCellLeftMargin);
            Right = new(ElementXml.TableCellRightMargin);

            if (ElementXml.TopMargin == null)
                ElementXml.TopMargin = Top._value;
            if (ElementXml.BottomMargin == null)
                ElementXml.BottomMargin = Bottom._value;
            if (ElementXml.TableCellLeftMargin == null)
                ElementXml.TableCellLeftMargin = Left._value;
            if (ElementXml.TableCellRightMargin == null)
                ElementXml.TableCellRightMargin = Right._value;
        }


        public static implicit operator OpenXmlElement.TableCellMarginDefault(TableMargin build) => build.ElementXml;
    }
}
