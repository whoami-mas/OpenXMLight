using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLight.validations
{
    public static class ValidationDocument
    {
        public static void ValidationWord(WordprocessingDocument? doc)
        {
            if (doc == null)
                throw new Exception("Документ не создан или неопределен");
        }

        public static void ValidateWidth(string width)
        {
            if (!int.TryParse(width, out int integerWidth))
                throw new ArgumentException("Не удалось преобразовать данное значение в ширину");
        }
    }
}
