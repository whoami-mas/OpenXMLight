using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXml = DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlF = DocumentFormat.OpenXml;

namespace OpenXMLight.Configurations.Elements
{
    public class Endnote
    {
        private string content;

        public TextProperties Properties { get; set; }
        public int ID { get; private set; }
        public string IdStyle { get; private set; }
        public string Content
        {
            get => content;
            set
            {
                content = value;

                Properties.Run.RemoveAllChildren<OpenXml.Text>();
                Properties.Run.Append(new OpenXml.Text(content));
            }
        }

        internal Endnote(string content, string idStyle, TextProperties? textProp = default)
        {
            this.IdStyle = idStyle;

            Create(textProp);

            this.Content = content;
        }

        internal void Create(TextProperties? textProp = default)
        {
            this.Properties = textProp ?? new();

            this.Properties.Paragraph.Append(
                new OpenXml.Run(
                    new OpenXml.RunProperties(
                        new OpenXml.RunStyle() { Val = IdStyle }
                    ),
                    new OpenXml.EndnoteReferenceMark()
                ),
                new OpenXml.Run(
                        new OpenXml.Text(" ") { Space = OpenXmlF.SpaceProcessingModeValues.Preserve }
                )
            );
            this.Properties.Paragraph.AppendChild(this.Properties.Run);
        }
        internal void SetID(int id) => this.ID = id;
    }
}
