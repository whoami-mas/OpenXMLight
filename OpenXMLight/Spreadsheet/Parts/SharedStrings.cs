using OpenXMLight.Configurations.Parts.InterfacesParts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLight.Spreadsheet.Parts
{
    internal class SharedStrings : IElementPart<OpenXmlPackaging.SharedStringTablePart>
    {
        public OpenXmlPackaging.SharedStringTablePart PartXml { get; set; }


        internal SharedStrings(OpenXmlPackaging.SharedStringTablePart sharedStringPart)
        {
            PartXml = sharedStringPart;

            CheckedExists();
        }

        public void CheckedExists() => PartXml.SharedStringTable ??= new OpenXmlSpreadsheet.SharedStringTable();



        public int AppendValue(string text)
        {
            OpenXmlSpreadsheet.SharedStringItem item = new OpenXmlSpreadsheet.SharedStringItem(new OpenXmlSpreadsheet.Text(text));

            PartXml.SharedStringTable.AppendChild<OpenXmlSpreadsheet.SharedStringItem>(item);

            return PartXml.SharedStringTable.ToList().IndexOf(item);
        }
        public string? GetValueOfIndex(int index)
        {
            OpenXmlSpreadsheet.SharedStringItem item = PartXml.SharedStringTable.ChildElements.OfType<OpenXmlSpreadsheet.SharedStringItem>()
                                                                                                .ToArray()[index];

            return item.Text?.Text;
        }
        public void RemoveElementOfIndex(int index)
        {
            PartXml.SharedStringTable.ChildElements.OfType<OpenXmlSpreadsheet.SharedStringItem>().ToArray()[index].Remove();
        }
    }
}
