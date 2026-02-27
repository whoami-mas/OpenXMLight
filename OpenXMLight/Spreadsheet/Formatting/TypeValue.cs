using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXMLight.config;

namespace OpenXMLight.Spreadsheet.Formatting
{
    public readonly record struct TypeValue : IEnumValue<int>
    {
        public int Value => _value;
        

        public static TypeValue General => new TypeValue(0);
        public static TypeValue Percent => new TypeValue(9);
        public static TypeValue Date => new TypeValue(14);
        public static TypeValue Number => new TypeValue(3);
        public static TypeValue Other => new TypeValue();



        private readonly int _value = 0;


        public TypeValue(int value)
        {
            _value = value;
        }


        public static TypeValue Parse(uint idFrmt)
        {
            //General formatte
            if (idFrmt == 0)
                return TypeValue.General;

            if(idFrmt == 9 || idFrmt == 10)
                return TypeValue.Percent;

            //Date formatte
            if (idFrmt >= 14 && idFrmt <= 17 || idFrmt == 22)
                return TypeValue.Date;

            if (idFrmt >= 1 && idFrmt <= 4 || idFrmt >= 37 && idFrmt <= 40)
                return TypeValue.Number;

            return TypeValue.Other;
        }
    }
}
