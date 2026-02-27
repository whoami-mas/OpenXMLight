using OpenXMLight.Configurations.Parts.InterfacesParts;
using OpenXMLight.Spreadsheet.Formatting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLight.Spreadsheet.Parts
{
    internal class Styles : IElementPart<OpenXmlPackaging.WorkbookStylesPart>
    {
        public OpenXmlPackaging.WorkbookStylesPart PartXml { get; set; }


        internal Styles(OpenXmlPackaging.WorkbookStylesPart stylePart)
        {
            PartXml = stylePart;

            CheckedExists();
        }


        public void CheckedExists() => PartXml.Stylesheet ??= new OpenXmlSpreadsheet.Stylesheet();


        public TypeValue GetFormatteCell(int indexStyle)
        {
            if (PartXml.Stylesheet.CellFormats == null)
                throw new Exception("Пустой список форматов ячеек");

            var format = PartXml.Stylesheet.CellFormats.OfType<OpenXmlSpreadsheet.CellFormat>().ToList()[indexStyle];

            return TypeValue.Parse(format.NumberFormatId);
        }

        public int AddStyle(TypeValue typeFormatte)
        {
            PartXml.Stylesheet.CellFormats ??= new OpenXmlSpreadsheet.CellFormats();

            OpenXmlSpreadsheet.CellFormat format = new()
            {
                NumberFormatId = Convert.ToUInt32(typeFormatte.Value),
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0
            };

            return PartXml.Stylesheet.CellFormats.OfType<OpenXmlSpreadsheet.CellFormat>().ToList().IndexOf(format);
        }
        public int GetFormatteCellIndex(TypeValue typeFormatte)
        {
            if (PartXml.Stylesheet.CellFormats == null)
                return AddStyle(typeFormatte);

            OpenXmlSpreadsheet.CellFormat cellFormatte =
                PartXml.Stylesheet.CellFormats.OfType<OpenXmlSpreadsheet.CellFormat>().FirstOrDefault(f => f.NumberFormatId == typeFormatte.Value);

            if (cellFormatte == null)
                return AddStyle(typeFormatte);

            return PartXml.Stylesheet.CellFormats.OfType<OpenXmlSpreadsheet.CellFormat>().ToList().IndexOf(cellFormatte);
        }
    }
}
