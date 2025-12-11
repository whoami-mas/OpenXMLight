using OpenXMLight.config;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.Configurations.Formatting
{
    public readonly record struct FontsFamily : IEnumValue<string>
    {
        public string Value => _value ?? "Arial";

        public static FontsFamily Arial => new FontsFamily("Arial");
        public static FontsFamily TimesNewRoman => new FontsFamily("Times New Roman");
        public static FontsFamily Calibri => new FontsFamily("Calibri");
        public static FontsFamily Verdana => new FontsFamily("Verdana");
        public static FontsFamily Tahoma => new FontsFamily("Tahoma");
        public static FontsFamily CourierNew => new FontsFamily("Courier New");
        public static FontsFamily Georgia => new FontsFamily("Georgia");
        public static FontsFamily PalatinoLinotype => new FontsFamily("Palatino Linotype");
        public static FontsFamily Garamond => new FontsFamily("Garamond");
        public static FontsFamily TrebuchetMS => new FontsFamily("Trebuchet MS");
        public static FontsFamily ComicSansMS => new FontsFamily("Comic Sans MS");
        public static FontsFamily LucidaConsole => new FontsFamily("Lucida Console");
        public static FontsFamily Consolas => new FontsFamily("Consolas");
        public static FontsFamily Cambria => new FontsFamily("Cambria");
        public static FontsFamily GillSans => new FontsFamily("Gill Sans");
        public static FontsFamily Impact => new FontsFamily("Impact");
        public static FontsFamily KristenITC => new FontsFamily("Kristen ITC");
        public static FontsFamily LucidaSansUnicode => new FontsFamily("Lucida Sans Unicode");
        public static FontsFamily CenturyGothic => new FontsFamily("Century Gothic");
        public static FontsFamily FranklinGothicMedium => new FontsFamily("Franklin Gothic Medium");


        private readonly string? _value;
        public FontsFamily(string fontFamily)
        {
            _value = fontFamily;
        }

        internal static FontsFamily Parse(string value)
        {
            switch (value)
            {
                case "Arial":
                    return Arial;
                case "Times New Roman":
                    return TimesNewRoman;
                case "Calibri":
                    return Calibri;
                case "Verdana":
                    return Verdana;
                case "Tahoma":
                    return Tahoma;
                case "Courier New":
                    return CourierNew;
                case "Georgia":
                    return Georgia;
                case "Palatino Linotype":
                    return PalatinoLinotype;
                case "Garamond":
                    return Garamond;
                case "Trebuchet MS":
                    return TrebuchetMS;
                case "Comic Sans MS":
                    return ComicSansMS;
                case "Lucida Console":
                    return LucidaConsole;
                case "Consolas":
                    return Consolas;
                case "Cambria":
                    return Cambria;
                case "Gill Sans":
                    return GillSans;
                case "Impact":
                    return Impact;
                case "Kristen ITC":
                    return KristenITC;
                case "Lucida Sans Unicode":
                    return LucidaSansUnicode;
                case "Century Gothic":
                    return CenturyGothic;
                case "Franklin Gothic Medium":
                    return FranklinGothicMedium;
                default:
                    return Arial;

            }
        }
    }
}
