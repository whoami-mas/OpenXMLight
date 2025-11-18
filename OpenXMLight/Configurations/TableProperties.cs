using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXMLight.Configurations.Formatting;
using OpenXML = DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlType = DocumentFormat.OpenXml;

namespace OpenXMLight.Configurations
{
    public class TableProperties
    {
        internal OpenXML.TableProperties TblPropXml { get; set; }

        private uint Size { get;}
        private BordersType Border { get; set; }
        private int MarginCell { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sizeBorder">size borders pt = (size / 8)</param>
        /// <param name="borderStyle"></param>
        /// <param name="paddingLR">pt margin cells left\right</param>
        public TableProperties(uint sizeBorder = 4, BordersType borderStyle = default, int paddingLR = 7)
        {
            this.Size = sizeBorder;
            this.Border = borderStyle;
            this.MarginCell = paddingLR * 15;

            this.Create();
        }

        private void Create()
        {
            TblPropXml = new OpenXML.TableProperties(
                new OpenXML.TableWidth() { Type = OpenXML.TableWidthUnitValues.Auto, Width = "0" },
                
                new OpenXML.TableBorders(
                    new OpenXML.TopBorder() { Val = Border.Value, Size = Size },
                    new OpenXML.BottomBorder() { Val = Border.Value, Size = Size },
                    new OpenXML.LeftBorder() { Val = Border.Value, Size = Size },
                    new OpenXML.RightBorder() { Val = Border.Value, Size = Size },
                    new OpenXML.InsideHorizontalBorder() { Val = Border.Value, Size = Size },
                    new OpenXML.InsideVerticalBorder() { Val = Border.Value, Size = Size }
                    ),
                new OpenXML.TableCellMarginDefault(
                    new OpenXML.TableCellLeftMargin() { Width = (OpenXmlType.Int16Value)MarginCell, Type = OpenXML.TableWidthValues.Dxa },
                    new OpenXML.TableCellRightMargin() { Width = (OpenXmlType.Int16Value)MarginCell, Type = OpenXML.TableWidthValues.Dxa }
                )
            );
        }
    }
}
