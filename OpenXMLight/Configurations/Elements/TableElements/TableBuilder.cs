using OpenXMLight.Configurations.Elements.TableElements;
using OpenXMLight.Configurations.Elements.TableElements.Formattings;
using OpenXMLight.Configurations.Elements.TableElements.Formattings.WidthComponents;
using OpenXMLight.Configurations.Elements.TableElements.Models;
using OpenXMLight.Configurations.Formatting;
using OpenXMLight.Configurations.Parts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXml = DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlFormatte = DocumentFormat.OpenXml;

namespace OpenXMLight.Configurations.Elements
{
    public class TableBuilder
    {
        private readonly Table table;
        private readonly SettingsPageWord? settingsDocument;


        //public TableBuilderTest() : this(new OpenXml.Table())
        //{

        //}
        internal TableBuilder(OpenXml.Table tbl, SettingsPageWord? settingsDocument = null)
        {
            table = new(tbl);
            this.settingsDocument = settingsDocument; 
        }



        public TableBuilder AddRows(Action<RowBuilder>? configuration = null)
        {
            var r = new OpenXml.TableRow();
            table.ElementXml.AppendChild(r);

            var builder = new RowBuilder(r);
            configuration?.Invoke(builder);

            if(
               table.Width.Type != TypeWidthTable.Auto ||
               !string.Equals(table.Width.Width, "0")
               )
                UpdateWidthCells((Row)builder);

            return this;
        }
        public TableBuilder SetBorders(Action<BordersLine>? configuration = null)
        {
            BordersLine borders = new BordersLine();
            
            configuration?.Invoke(borders);

            table.Borders = borders;
            
            return this;
        }
        public TableBuilder SetWidth(Action<TableCellWidth<TableWidth>>? configuration = null)
        {
            TableCellWidth<TableWidth> width = new();
            configuration?.Invoke(width);

            table.Width = width;

            return this;
        }
        public TableBuilder IsFixed(bool isFixed)
        {
            table.IsFixed = isFixed;

            return this;
        }
        public TableBuilder Merge(int c1_r, int c1_c,
            int c2_r, int c2_c)
        {
            if (table.Rows.Count < 1 || table.CountColumn < 0)
                throw new ArgumentException("Нет ячеек для соединения");

            table.MergeCell(c1_r, c1_c, c2_r, c2_c);

            return this;
        }
        public TableBuilder SetMargin(string left, string top, string right, string bottom)
        {
            if (!double.TryParse(left, out double dLeft) ||
                !double.TryParse(top, out double dTop) ||
                !double.TryParse(right, out double dRight) ||
                !double.TryParse(bottom, out double dBottom))
                throw new FormatException("Некорректный формат отступов");
            
            if (dLeft < 0 || dTop < 0 || dRight < 0 || dBottom < 0)
                throw new ArgumentException("Отступы не могут быть отрицательными значениями");


            table.Margin.Left.Width = dLeft.ToString();
            table.Margin.Top.Width = dTop.ToString();
            table.Margin.Right.Width = dRight.ToString();
            table.Margin.Bottom.Width = dBottom.ToString();

            return this;
        }


        public static implicit operator Table(TableBuilder build) => build.table;


        private void UpdateWidthCells(Row row)
        {
            if (double.TryParse(table.Width.Width, out double width))
            {
                double widthCell = width / row.Cells.Count;

                foreach(Cell cell in row.Cells)
                {
                    cell.Width = new();
                    cell.Width.Width = cell.Merged > 1 
                        ? (widthCell * cell.Merged).ToString()
                        : widthCell.ToString();
                    cell.Width.Type = table.Width.Type;
                }
            }
            else
                throw new ArgumentException("Не получилось преобразовать ширину таблицы");
        }
    }
}
