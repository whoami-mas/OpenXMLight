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

        public Text(string content, TextProperties? textProp = default)
        {
            Create(textProp);

            this.Content = content;
        }

        internal void Create(TextProperties? textProp = default)
        {
            Properties ??= new();

            Properties.Paragraph.AppendChild(Properties.Run);
        }
    }
}
