using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXML = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Elements
{
    public class Text
    {
        private string content;
        private Endnote endnote;

        public TextProperties Properties { get; set; }
        public string Content
        {
            get => content;
            set
            {
                content = value;

                Properties.Run.RemoveAllChildren<OpenXML.Text>();
                Properties.Run.Append(new OpenXML.Text(content));
            }
        }
        public Endnote? Endnote {
            get => endnote;
            set 
            {
                if(value != null)
                {
                    endnote = value;


                    Properties.Paragraph.RemoveChild(Properties.Paragraph.Elements<OpenXML.EndnoteReference>().FirstOrDefault()?.Parent);
                    Properties.Paragraph.AppendChild(
                        new OpenXML.Run(
                            new OpenXML.RunProperties(
                                new OpenXML.RunStyle() { Val = endnote?.IdStyle }
                            ),
                            new OpenXML.EndnoteReference() { Id = endnote?.ID }
                        )
                    );
                }
            }
        }

        public Text(string content, Endnote? endnote = default, TextProperties? textProp = default)
        {
            Create(textProp);

            this.Content = content;
            this.Endnote = endnote;
        }

        internal void Create(TextProperties? textProp = default)
        {
            this.Properties = textProp ?? new();

            this.Properties.Paragraph.AppendChild(this.Properties.Run);
        }
    }
}
