using OpenXMLight.Configurations.Elements.TableElements.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Elements.TableElements
{
    public class RowBuilder
    {
        private readonly Row _row;



        public RowBuilder() : this(new OpenXml.TableRow())
        {

        }
        internal RowBuilder(OpenXml.TableRow r) => _row = new Row(r);



        public RowBuilder AddCell(Action<CellBuilder>? configuration = null)
        {
            OpenXml.TableCell appendCell = new OpenXml.TableCell();
            _row.ElementXml.AppendChild(appendCell);

            var cellBuilder = new CellBuilder(appendCell);

            configuration?.Invoke(cellBuilder);

            return this;
        }



        public static implicit operator Row(RowBuilder build) => build._row;
    }
}
