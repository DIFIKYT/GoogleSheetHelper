using Google.Apis.Sheets.v4.Data;
using System.Text;

namespace GoogleSheetsHelper
{
    public static class OtherMethods
    {
        public static int GetColumnIndex(string columnLetter)
        {
            int columnIndex = 0;
            columnLetter = columnLetter.ToUpper();

            for (int i = 0; i < columnLetter.Length; i++)
            {
                char letter = columnLetter[i];
                columnIndex = columnIndex * 26 + (letter - 'A' + 1);
            }

            return columnIndex;
        }

        public static string GetColumnName(int columnIndex)
        {
            StringBuilder columnName = new();

            while (columnIndex > 0)
            {
                int remainder = (columnIndex - 1) % 26;
                char letter = (char)('A' + remainder);
                columnName.Insert(0, letter);
                columnIndex = (columnIndex - 1) / 26;
            }

            return columnName.ToString();
        }

        public static bool TryGetList(string listName, Spreadsheet sourceSpreadsheet)
        {
            try
            {
                if (string.IsNullOrEmpty(listName))
                    return false;

                return sourceSpreadsheet.Sheets.Any(sheet => sheet.Properties.Title == listName);
            }
            catch (Exception exception)
            {
                Console.WriteLine($"\n\nОшибка в TryGetList '{sourceSpreadsheet.Properties.Title}'\n{exception}\n\n");
                return false;
            }
        }
    }
}