using OpenXMLight.Configurations.Elements.Interfaces;
using OpenXMLight.Configurations.Elements.TableElements.Formattings.FormattingsBase;
using OpenXMLight.Configurations.Formatting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlEl = DocumentFormat.OpenXml;

namespace OpenXMLight.Configurations.Elements.TableElements.Formattings.WidthComponents
{
    public class TableWidth : WidthBase, IElementBase<OpenXml.TableWidth>
    {
        private string width = "0";
        private TypeWidthTable type = TypeWidthTable.Auto;


        public OpenXml.TableWidth? ElementXml { get; set; }
        public override string Width
        {
            get => width;
            set
            {
                width = value;

                ElementXml.Width = GetWidthOfType<string>(width, Type);
            }
        }
        public override TypeWidthTable Type
        {
            get => type;
            set
            {
                if (!IsTypeFirst)
                {
                    type = value;
                    ElementXml.Width = GetWidthOfType<string>(Width, Type);
                    IsTypeFirst = true;
                }
                else
                {

                }

                ElementXml.Type = type.Value;
            }
        }


        internal override void Initialize(OpenXmlEl.OpenXmlElement? xml)
        {
            if (xml == null)
                ElementXml = new();
            else
            {
                ElementXml = xml as OpenXml.TableWidth;
                type = TypeWidthTable.Parse(ElementXml.Type);
                width = ConvertWidth(ElementXml.Width, Type);
            }
        }



        public static implicit operator OpenXml.TableWidth(TableWidth build) => build.ElementXml;
    }
}
