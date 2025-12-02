using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OpenXMLight.Tools
{
    internal static class HalperData
    {
        internal static int GetRowIndex(string input)
        {
            Match regexMatch = Regex.Match(input, @"\d+", RegexOptions.IgnoreCase);
            if (regexMatch.Success)
            {
                return int.Parse(regexMatch.Value);
            }
            else
                return 0 ;
        }
        internal static int GetColumnIndex(string input)
        {
            int index = 0;
            string column = Regex.Match(input, @"[A-Z]+", RegexOptions.IgnoreCase).Value;

            for (int i = 0; i < column.Length; i++)
            {
                index *= 26;
                index += (column[i] - 'A' + 1);
            }
            return index;
        }
        internal static string GetColumnByIndex(int index)
        {
            string columnName = string.Empty;
            while (index > 0)
            {
                int remainder = (index - 1) % 26;
                columnName = (char)(remainder + 'A') + columnName;
                index = (index - 1) / 26;
            }
            return columnName;
        }
    }
}
