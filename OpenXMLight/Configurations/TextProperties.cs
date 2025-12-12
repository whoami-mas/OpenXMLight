using OpenXMLight.Configurations.Formatting;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXML = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations
{
    public class TextProperties
    {
        internal OpenXML.Paragraph Paragraph { get; set; }
        internal OpenXML.Run Run { get; set; }


        private int fontSize = 11;
        private FontsFamily fontFamily = FontsFamily.Calibri;
        private bool bold = false;
        private HorizonatalAlignments hAlignment = HorizonatalAlignments.Left;
        private SpacingBetweenLines spBetLines;


        public int FontSize 
        {
            get => fontSize;
            set
            {
                fontSize = value;

                Run.RunProperties.FontSize = null;
                Run.RunProperties.FontSize = new OpenXML.FontSize() { Val = (fontSize * 2).ToString() };
            }
        }
        public FontsFamily FontFamily 
        {
            get => fontFamily;
            set
            {
                fontFamily = value;

                Run.RunProperties.RunFonts = null;
                Run.RunProperties.RunFonts = new OpenXML.RunFonts() { Ascii = fontFamily.Value, HighAnsi = fontFamily.Value };
            }
        }
        public bool Bold 
        {
            get => bold;
            set
            {
                bold = value;

                Run.RunProperties.Bold = bold ? Run.RunProperties.Bold = new OpenXML.Bold()
                                       : null;
            }
        }
        public HorizonatalAlignments HAlignment 
        {
            get => hAlignment;
            set
            {
                hAlignment = value;

                Paragraph.ParagraphProperties.Justification = null;
                Paragraph.ParagraphProperties.Justification = new OpenXML.Justification() { Val = hAlignment.Value };
            }
        }
        public SpacingBetweenLines SpBetLines 
        {
            get => spBetLines;
            set 
            {
                if (spBetLines != null)
                {
                    spBetLines.PropertyChanged -= SpBetLines_PropertyChanged;
                }
                spBetLines = value;
                if (spBetLines != null)
                {
                    spBetLines.PropertyChanged += SpBetLines_PropertyChanged;
                }

                Paragraph.ParagraphProperties.SpacingBetweenLines = null;
                Paragraph.ParagraphProperties.SpacingBetweenLines = new OpenXML.SpacingBetweenLines() { After = spBetLines.After.ToString(),
                                                                                                          Before = spBetLines.Before.ToString(),
                                                                                                          Line = spBetLines.Line.ToString(),
                                                                                                          LineRule = OpenXML.LineSpacingRuleValues.Auto};
            }
        }



        public TextProperties() => this.Create();

        internal TextProperties(TextProperties textProp)
        {
            this.Create();

            FontSize = textProp.FontSize;
            FontFamily = textProp.FontFamily;
            Bold = textProp.Bold;
            HAlignment = textProp.HAlignment;
            SpBetLines = textProp.SpBetLines;
        }
        internal TextProperties(OpenXML.Paragraph paragraph)
        {
            Create(paragraph);

            FontSize = Run.RunProperties.FontSize != null ? int.Parse(Run.RunProperties.FontSize.Val) / 2
                : 11;
            FontFamily = FontsFamily.Parse(Run.RunProperties.RunFonts?.HighAnsi.Value);
            Bold = Run.RunProperties.Bold != null ? true : false;
            HAlignment = HorizonatalAlignments.Parse(Paragraph.ParagraphProperties.Justification?.Val);
        }



        internal void Create(OpenXML.Paragraph? p = default)
        {
            Paragraph = p ?? new OpenXML.Paragraph();
            Paragraph.ParagraphProperties = p.ParagraphProperties != null ? p.ParagraphProperties : p.ParagraphProperties = new OpenXML.ParagraphProperties();


            Run = p?.Elements<OpenXML.Run>().FirstOrDefault() ?? p.AppendChild(new OpenXML.Run(
                new OpenXML.RunProperties()
                ));

            SpBetLines = new();
            SpBetLines.PropertyChanged += SpBetLines_PropertyChanged;
        }

        private void SpBetLines_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            Paragraph.ParagraphProperties.SpacingBetweenLines ??= new OpenXML.SpacingBetweenLines();

            if (sender is SpacingBetweenLines sp)
                switch (e.PropertyName)
                {
                    case nameof(SpacingBetweenLines.After):
                        Paragraph.ParagraphProperties.SpacingBetweenLines.After = sp.After.ToString();
                        break;
                    case nameof(SpacingBetweenLines.Before):
                        Paragraph.ParagraphProperties.SpacingBetweenLines.Before = sp.Before.ToString();
                        break;
                    case nameof(SpacingBetweenLines.Line):
                        Paragraph.ParagraphProperties.SpacingBetweenLines.Line = sp.Line.ToString();
                        break;
                }
        }
    }
}
