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
        private uint size;
        private BordersType borders;
        private int marginCell;

        internal OpenXML.TableProperties TblPropXml { get; set; }

        public uint Size { get => size; set 
            {
                size = value;

                ValidateBorders();

                TblPropXml.TableBorders.TopBorder.Size = value;
                TblPropXml.TableBorders.BottomBorder.Size = value;
                TblPropXml.TableBorders.LeftBorder.Size = value;
                TblPropXml.TableBorders.RightBorder.Size = value;
                TblPropXml.TableBorders.InsideHorizontalBorder.Size = value;
                TblPropXml.TableBorders.InsideVerticalBorder.Size = value;
            } }
        public BordersType Border { get => borders; set 
            {
                borders = value;

                ValidateBorders();

                TblPropXml.TableBorders.TopBorder.Val = value.Value;
                TblPropXml.TableBorders.BottomBorder.Val = value.Value;
                TblPropXml.TableBorders.LeftBorder.Val = value.Value;
                TblPropXml.TableBorders.RightBorder.Val = value.Value;
                TblPropXml.TableBorders.InsideHorizontalBorder.Val = value.Value;
                TblPropXml.TableBorders.InsideVerticalBorder.Val = value.Value;
            } }
        public int MarginCell { get => marginCell / Configuration.TwipsInPixels;
            set {
                marginCell = value * Configuration.TwipsInPixels;

                TblPropXml.TableCellMarginDefault.TableCellLeftMargin.Width = Convert.ToInt16(marginCell);
                TblPropXml.TableCellMarginDefault.TableCellRightMargin.Width = Convert.ToInt16(marginCell);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sizeBorder">size borders pt = (size / 8)</param>
        /// <param name="borderStyle"></param>
        /// <param name="paddingLR">pt margin cells left\right</param>
        public TableProperties()
        {
            this.Create();

            //this.Size = sizeBorder;
            //this.Border = borderStyle;
            //this.MarginCell = paddingLR * Configuration.TwipsInPixels;
        }

        private void Create()
        {
            TblPropXml = new OpenXML.TableProperties(
                new OpenXML.TableWidth() { Type = OpenXML.TableWidthUnitValues.Auto, Width = "0" },
                
                //new OpenXML.TableBorders(
                //    new OpenXML.TopBorder() { Val = Border.Value, Size = Size },
                //    new OpenXML.BottomBorder() { Val = Border.Value, Size = Size },
                //    new OpenXML.LeftBorder() { Val = Border.Value, Size = Size },
                //    new OpenXML.RightBorder() { Val = Border.Value, Size = Size },
                //    new OpenXML.InsideHorizontalBorder() { Val = Border.Value, Size = Size },
                //    new OpenXML.InsideVerticalBorder() { Val = Border.Value, Size = Size }
                //    ),
                new OpenXML.TableCellMarginDefault(
                    new OpenXML.TableCellLeftMargin() { Width = (OpenXmlType.Int16Value)MarginCell, Type = OpenXML.TableWidthValues.Dxa },
                    new OpenXML.TableCellRightMargin() { Width = (OpenXmlType.Int16Value)MarginCell, Type = OpenXML.TableWidthValues.Dxa }
                )
            );
        }

        private void ValidateBorders()
        {
            if (TblPropXml.TableBorders == null)
                TblPropXml.TableBorders = new(
                        new OpenXML.TopBorder(),
                        new OpenXML.BottomBorder(),
                        new OpenXML.LeftBorder(),
                        new OpenXML.RightBorder(),
                        new OpenXML.InsideHorizontalBorder(),
                        new OpenXML.InsideVerticalBorder()
                    );
        }
    }
}
