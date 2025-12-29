using OpenXMLight.Configurations.Elements.Interfaces;
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
    public class CellMargin : IMarginContainer, IElementBase<OpenXmlElement.TableCellMargin>
    {
        public OpenXmlElement.TableCellMargin ElementXml { get; set; }
        public Margin<TopMargin> Top { get; protected set; }
        public Margin<BottomMargin> Bottom { get; protected set; }
        public Margin<LeftMargin> Left { get; protected set; }
        public Margin<RightMargin> Right { get; protected set; }


        IMargin IMarginContainer.Top => Top;
        IMargin IMarginContainer.Bottom => Bottom;
        IMargin IMarginContainer.Left => Left;
        IMargin IMarginContainer.Right => Right;

        public void Initialize(OpenXml.OpenXmlElement? xml = null)
        {
            ElementXml = xml as OpenXmlElement.TableCellMargin ?? new();

            Top = new(ElementXml.TopMargin);
            Bottom = new(ElementXml.BottomMargin);
            Left = new(ElementXml.LeftMargin);
            Right = new(ElementXml.RightMargin);

            if (ElementXml.TopMargin == null)
                ElementXml.TopMargin = Top._value;
            if (ElementXml.BottomMargin == null)
                ElementXml.BottomMargin = Bottom._value;
            if (ElementXml.LeftMargin == null)
                ElementXml.LeftMargin = Left._value;
            if (ElementXml.RightMargin == null)
                ElementXml.RightMargin = Right._value;
        }


        public static implicit operator OpenXmlElement.TableCellMargin(CellMargin build) => build.ElementXml;
    }
}
