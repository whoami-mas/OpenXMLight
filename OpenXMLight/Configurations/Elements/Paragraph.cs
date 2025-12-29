using OpenXMLight.Configurations.Formatting;
using OpenXMLight.Tools;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml.Wordprocessing;
using values = DocumentFormat.OpenXml;

namespace OpenXMLight.Configurations.Elements
{
    public class Paragraph : Element<OpenXml.Paragraph, OpenXml.ParagraphProperties>
    {
        internal override OpenXml.Paragraph ElementXml { get; set; }
        internal override OpenXml.ParagraphProperties ElementProperties
        {
            get
            {
                if(_elementProperties == null)
                    _elementProperties = ElementXml.ParagraphProperties ??= new OpenXml.ParagraphProperties();

                return _elementProperties;
            }
        }


        //Constructor
        internal Paragraph(OpenXml.Paragraph p) => ElementXml = p;



        
        #region Private properties
        private OpenXml.ParagraphProperties? _elementProperties;
        private ElementCollection<Run>? _runs;
        private SpacingBetweenLines? _spacing;
        private ElementCollection<EndnoteTest> _endnote;
        #endregion

        //Properties
        public HorizontalAlignments Alignment
        {
            get
            {
                object? val = ElementXml.ParagraphProperties?.Justification?.Val;

                return HelperData.TryParseParagraphAlignment(val, out HorizontalAlignments alignment)
                    ? alignment
                    : Configuration.DEFAULT_HORIZONTAL_ALIGNMENT;
            }
            set => ElementProperties.Justification = new OpenXml.Justification() { Val = value.Value };
        }
        public SpacingBetweenLines Spacing 
        { 
            get
            {
                if(_spacing == null)
                {
                    OpenXml.SpacingBetweenLines? spacingXml = ElementXml.ParagraphProperties?.SpacingBetweenLines;

                    _spacing = HelperData.TryParseSpacingBetweenLines(spacingXml, out SpacingBetweenLines spacing)
                        ? spacing
                        : Configuration.DEFAULT_SPACING_BETWEEN_LINES;

                    if (_spacing is INotifyPropertyChanged)
                        _spacing.PropertyChanged += Spacing_PropertyChanged;
                }

                return _spacing;
            }
            set
            {
                if (_spacing == value)
                    return;

                _spacing = value;
            }
        }
        public string AllText => ElementXml.InnerText;
        public ElementCollection<Run> Runs
        {
            get
            {
                if(_runs == null)
                    _runs = new(this.ElementXml.Elements<OpenXml.Run>().Select(s => new Run(s))) { Parent = this.ElementXml };

                return _runs;
            }
        }
        public ElementCollection<EndnoteTest> Endnotes
        {
            get
            {
                if(_endnote == null)
                {
                    OpenXml.Run r = new();
                    

                    _endnote = new(ElementXml.Elements<OpenXml.Run>().Where(w=> w.Elements<OpenXml.EndnoteReference>().Count() > 0).Select(s=> new EndnoteTest(s)));
                }

                return _endnote;
            }
        }
       

        private void Spacing_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            ElementProperties.SpacingBetweenLines ??= new OpenXml.SpacingBetweenLines();

            if (sender is SpacingBetweenLines sp)
                switch (e.PropertyName)
                {
                    case nameof(SpacingBetweenLines.After):
                        ElementXml.ParagraphProperties.SpacingBetweenLines.After = (sp.After * Configuration.InchInPixels).ToString();
                        break;
                    case nameof(SpacingBetweenLines.Before):
                        ElementXml.ParagraphProperties.SpacingBetweenLines.Before = (sp.Before * Configuration.InchInPixels).ToString();
                        break;
                    case nameof(SpacingBetweenLines.Line):
                        ElementXml.ParagraphProperties.SpacingBetweenLines.Line = (sp.Line * Configuration.InchInPixels).ToString();
                        break;
                }
        }
    }
}
