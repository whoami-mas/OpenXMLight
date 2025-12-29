using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLight.config;
using OpenXMLight.Configurations.Elements.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.Configurations.Formatting
{
    public readonly record struct TypeWidthTable : IEnumValue<TableWidthUnitValues>
    {
        public TableWidthUnitValues Value => _value ?? TableWidthUnitValues.Dxa;


        public static TypeWidthTable Pct => new TypeWidthTable(TableWidthUnitValues.Pct);
        public static TypeWidthTable Auto => new TypeWidthTable(TableWidthUnitValues.Auto);
        public static TypeWidthTable Cm => new TypeWidthTable(TableWidthUnitValues.Dxa);


        private readonly TableWidthUnitValues? _value;
        public TypeWidthTable(TableWidthUnitValues value) => _value = value;


        internal static TypeWidthTable Parse(TableWidthUnitValues value)
        {
            return value switch
            {
                var v when v == TableWidthUnitValues.Auto => TypeWidthTable.Auto,
                var v when v == TableWidthUnitValues.Pct => TypeWidthTable.Pct,
                var v when v == TableWidthUnitValues.Dxa => TypeWidthTable.Cm
            };
        }
    }
}
