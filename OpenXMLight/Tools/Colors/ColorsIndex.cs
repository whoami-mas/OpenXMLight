using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXMLight.Configurations.Formatting;

namespace OpenXMLight.Tools.Colors
{
    internal static class ColorsIndex
    {
        private static Dictionary<uint, Color> _indexedColor = new()
        {
            { 0, Color.FromHex("#000000") }, // Черный
            { 1, Color.FromHex("#FFFFFF") }, // Белый
            { 2, Color.FromHex("#FF0000") }, // Красный
            { 3, Color.FromHex("#00FF00") }, // Зеленый
            { 4, Color.FromHex("#0000FF") }, // Синий
            { 5, Color.FromHex("#FFFF00") }, // Желтый
            { 6, Color.FromHex("#FF00FF") }, // Пурпурный
            { 7, Color.FromHex("#00FFFF") }, // Голубой

            { 8, Color.FromHex("#800000") }, // Темно-красный
            { 9, Color.FromHex("#008000") }, // Темно-зеленый
            { 10, Color.FromHex("#000080") }, // Темно-синий
            { 11, Color.FromHex("#808000") }, // Оливковый
            { 12, Color.FromHex("#800080") }, // Темно-пурпурный
            { 13, Color.FromHex("#008080") }, // Бирюзовый
            { 14, Color.FromHex("#C0C0C0") }, // Серебряный
            { 15, Color.FromHex("#808080") }, // Серый

            { 16, Color.FromHex("#9999FF") }, // Сиреневый
            { 17, Color.FromHex("#993366") }, // Розово-коричневый
            { 18, Color.FromHex("#FFFFCC") }, // Светло-желтый
            { 19, Color.FromHex("#CCFFFF") }, // Светло-голубой
            { 20, Color.FromHex("#660066") }, // Фиолетовый
            { 21, Color.FromHex("#FF8080") }, // Светло-красный
            { 22, Color.FromHex("#0066CC") }, // Синий (темный)
            { 23, Color.FromHex("#CCCCFF") }, // Светло-сиреневый
            { 24, Color.FromHex("#000080") }, // Темно-синий (Navy)
            { 25, Color.FromHex("#FF00FF") }, // Розовый
            { 26, Color.FromHex("#FFFF00") }, // Ярко-желтый
            { 27, Color.FromHex("#00FFFF") }, // Ярко-голубой
            { 28, Color.FromHex("#800080") }, // Фиолетовый (темный)
            { 29, Color.FromHex("#800000") }, // Бордовый
            { 30, Color.FromHex("#008080") }, // Зеленовато-голубой
            { 31, Color.FromHex("#0000FF") }, // Ярко-синий

            { 32, Color.FromHex("#00CCFF") }, // Небесно-голубой
            { 33, Color.FromHex("#CCFFFF") }, // Бледно-голубой
            { 34, Color.FromHex("#CCFFCC") }, // Светло-зеленый
            { 35, Color.FromHex("#FFFF99") }, // Светло-желтый
            { 36, Color.FromHex("#99CCFF") }, // Светло-синий
            { 37, Color.FromHex("#FF99CC") }, // Светло-розовый
            { 38, Color.FromHex("#CC99FF") }, // Светло-фиолетовый
            { 39, Color.FromHex("#FFCC99") }, // Персиковый
            { 40, Color.FromHex("#3366FF") }, // Ярко-синий 2
            { 41, Color.FromHex("#33CCCC") }, // Бирюзовый
            { 42, Color.FromHex("#99CC00") }, // Ярко-зеленый
            { 43, Color.FromHex("#FF99CC") }, // Розовый (светлый)
            { 44, Color.FromHex("#FFCC99") }, // Оранжевый (светлый)
            { 45, Color.FromHex("#FFFF99") }, // Кремовый
            { 46, Color.FromHex("#CCFFCC") }, // Мятный
            { 47, Color.FromHex("#CCFFFF") }, // Бледно-голубой 2

            { 48, Color.FromHex("#99CCFF") }, // Голубой
            { 49, Color.FromHex("#CC99FF") }, // Сиреневый
            { 50, Color.FromHex("#FF99CC") }, // Розовый 2
            { 51, Color.FromHex("#FFCC99") }, // Абрикосовый
            { 52, Color.FromHex("#FFFF99") }, // Желтый (светлый) 2
            { 53, Color.FromHex("#CCFFCC") }, // Зеленый (светлый) 2
            { 54, Color.FromHex("#CCFFFF") }, // Голубой (светлый) 2
            { 55, Color.FromHex("#99CCFF") },  // Синий (светлый) 3
            { 64, Color.FromHex("#000000") } // Автоматический
        };

        public static Dictionary<uint, Color> indexedColor => _indexedColor;
    }
}
