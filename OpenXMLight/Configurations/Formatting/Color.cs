using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.Configurations.Formatting
{
    public readonly record struct Color
    {
        public string Hex => _hex;
        public byte R => GetComponent(0);
        public byte G => GetComponent(1);
        public byte B => GetComponent(2);
        public byte A => GetComponent(-1);


        private readonly string _hex;
        public Color(string hex)
        {
            ValidationHex(hex);

            _hex = hex;
        }



        public static Color FromHex(string hex)
        {
            ValidationHex(hex);

            return new(hex);
        }
        public static Color FromRgb(byte red, byte green, byte blue) =>
            new Color($"#{red:X2}{green:X2}{blue:X2}");

        public static Color FromArgb(byte alpha, byte red, byte green, byte blue) =>
            new Color($"#{alpha:X2}{red:X2}{green:X2}{blue:X2}");



        private static void ValidationHex(string hexCode)
        {
            if (string.IsNullOrWhiteSpace(hexCode))
                throw new ArgumentException("HEX код не можен быть пустым");

            if (!hexCode.StartsWith('#'))
                throw new ArgumentException("Не правильный формат HEX кода");

            if (hexCode.Length != 4 && hexCode.Length != 7 && hexCode.Length != 9)
                throw new ArgumentException("Не правильная длина HEX кода");
        }
        private byte GetComponent(int index)
        {
            if (index == -1)
                if (_hex.Length == 9)
                    return Convert.ToByte(_hex.Substring(1, 2), 16);
                else
                    return 255;

            if (_hex.Length == 4)
            {
                var s = _hex[index];
                return Convert.ToByte($"{s}{s}", 16);
            }

            int startPosition = _hex.Length == 9 ? 3 : 1;
            int position = startPosition + (index * 2);

            return Convert.ToByte(_hex.Substring(position, 2), 16);
        }
    }
}
