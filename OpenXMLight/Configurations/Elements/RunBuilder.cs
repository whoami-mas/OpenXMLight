using OpenXMLight.Configurations.Formatting;
using OpenXMLight.Configurations.WordContext;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlFormatte = DocumentFormat.OpenXml;

namespace OpenXMLight.Configurations.Elements
{
    public class RunBuilder
    {
        Run _run;



        public RunBuilder() => _run = new();
        internal RunBuilder(OpenXml.Run r) => _run = new(r);



        public RunBuilder SetText(string text)
        {
            var t = _run.ElementXml.Elements<OpenXml.Text>().FirstOrDefault();

            if (t == null)
                _run.ElementXml.AppendChild(new OpenXml.Text(text) { Space = text.StartsWith(' ') || text.EndsWith(' ') 
                    ? OpenXmlFormatte.SpaceProcessingModeValues.Preserve 
                    : null});
            else
            {
                t.Text = text;
                t.Space = text.StartsWith(' ') || text.EndsWith(' ')
                    ? OpenXmlFormatte.SpaceProcessingModeValues.Preserve
                    : null;
            }

            return this;
        }
        public RunBuilder SetFontSize(int fontSize)
        {
            _run.ElementProperties.FontSize ??= new OpenXml.FontSize();
            _run.ElementProperties.FontSize.Val = (fontSize * 2).ToString();

            return this;
        }
        public RunBuilder SetFontFamily(FontsFamily fontFamily)
        {
            _run.ElementProperties.RunFonts ??= new OpenXml.RunFonts();
            _run.ElementProperties.RunFonts.Ascii = fontFamily.Value;
            _run.ElementProperties.RunFonts.HighAnsi = fontFamily.Value;

            return this;
        }
        public RunBuilder SetBold(bool bold)
        {
            _run.ElementProperties.Bold = bold ? new OpenXml.Bold() 
                                               : null;

            return this;
        }
        public RunBuilder SetEndnote(Endnote endnote)
        {
            _run.ElementProperties.AppendChild(
                new OpenXml.RunStyle() { Val = Context.GetInstance().Styles.CreateGetEndnoteRef() }
                );
            _run.ElementXml.AppendChild(
                new OpenXml.EndnoteReference() { Id = endnote.ElementXml.Id }
                );

            return this;
        }
        public RunBuilder SetColor(Color color)
        {
            _run.Color = color;

            return this;
        }



        public static implicit operator Run(RunBuilder build)
        {
            return build._run;
        }
    }
}
