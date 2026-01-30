using DocumentFormat.OpenXml;
using OpenXMLight.Configurations.Formatting;
using OpenXMLight.Tools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlFormatte = DocumentFormat.OpenXml;

namespace OpenXMLight.Configurations.Elements
{
    public class Run : Element<OpenXml.Run, OpenXml.RunProperties>
    {
        internal override OpenXml.Run ElementXml { get; set; }
        internal override OpenXml.RunProperties ElementProperties
        { 
            get 
            {
                if(_elementProperties == null)
                    _elementProperties = ElementXml.RunProperties ??= new OpenXml.RunProperties();

                return _elementProperties;
            } 
        }


        internal Run() => ElementXml = new OpenXml.Run();
        internal Run(OpenXml.Run r) => ElementXml = r;



        #region Private properties
        private OpenXml.RunProperties? _elementProperties;
        private Color? _color;
        #endregion

        public string? Text { 
            get => ElementXml.Elements<OpenXml.Text>().FirstOrDefault()?.Text;
            set
            {
                ElementXml.RemoveAllChildren();

                var text = new OpenXml.Text(value);
                ElementXml.AppendChild(new OpenXml.Text(value)
                {
                    Space = value.StartsWith(' ') || value.EndsWith(' ')
                    ? OpenXmlFormatte.SpaceProcessingModeValues.Preserve
                    : null
                });
            }
        }
        public int FontSize {
            get 
            {
                string? fontSizeVal = ElementXml.RunProperties?.FontSize?.Val;

                return HelperData.TryParseFontSize(fontSizeVal, out int fontSize)
                    ? fontSize
                    : Configuration.DEFAULT_FONTSIZE;
            }
            set => ElementProperties.FontSize = new OpenXml.FontSize() { Val = (value * 2).ToString() };
        }
        public FontsFamily FontFamily {
            get
            {
                string? fontFamilyValue = ElementXml.RunProperties?.RunFonts?.Ascii;

                return HelperData.TryParseFontFamily(fontFamilyValue, out FontsFamily fontFamily)
                    ? fontFamily
                    : Configuration.DEFAULT_FONTFAMILY;
            }
            set => ElementProperties.RunFonts = new OpenXml.RunFonts() { Ascii = value.Value, HighAnsi = value.Value };
        }
        public bool Bold {
            get 
            {
                var bold = ElementXml.RunProperties?.Bold;

                return bold != null ? true : false;
            }
            set => ElementProperties.Bold = value ? new OpenXml.Bold()
                                                       : null;
        }
        public Color? Color
        {
            get
            {
                object? tmpColor = ElementProperties.Color?.Val;

                _color = HelperData.TryParseColorText(tmpColor, out Color? _result)
                    ? _result
                    : Configuration.DEFAULT_COLOR_TEXT;

                return _color;
            }
            set
            {
                if (Color == value)
                    return;

                _color = value;

                if(_color == null)
                {
                    ElementProperties.Color.Remove();
                    return;
                }

                ElementProperties.Color ??= new OpenXml.Color();
                ElementProperties.Color.Val = _color.Value.Hex.Substring(1);
            }
        }
    }
}
