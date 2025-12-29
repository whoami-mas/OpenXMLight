using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLight.Configurations.Elements.TableElements.Models
{
    public class Row : Element<OpenXml.TableRow, OpenXml.TableRowProperties>
    {
        internal override OpenXml.TableRow ElementXml { get; set; }
        internal override OpenXml.TableRowProperties ElementProperties
        {
            get
            {
                if (_elementProperties == null)
                    _elementProperties = ElementXml.TableRowProperties ??= new OpenXml.TableRowProperties();

                return _elementProperties;
            }
        }



        internal Row(OpenXml.TableRow r) => ElementXml = r;



        #region Private properties
        private OpenXml.TableRowProperties? _elementProperties;
        private ElementCollection<Cell>? _cells;
        #endregion

        public ElementCollection<Cell> Cells
        {
            get
            {
                if (_cells == null)
                    _cells = new(ElementXml.Elements<OpenXml.TableCell>().Select(s => new Cell(s))) { Parent = ElementXml };

                return _cells;
            }
        }
    }
}
