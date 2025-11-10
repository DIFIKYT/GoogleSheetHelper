namespace GoogleSheetsHelper
{
    public class StartEndRange(int startRowNumber, int endRowNumber, string startColumnLetter, string endColumnLetter)
    {
        public int StartRowNumber { get; init; } = startRowNumber;
        public int EndRowNumber { get; init; } = endRowNumber;
        public string StartColumnLetter { get; init; } = startColumnLetter;
        public string EndColumnLetter { get; init; } = endColumnLetter;

        public int StartRowIndex { get; init; } = startRowNumber - 1;
        public int EndRowIndex { get; init; } = endRowNumber - 1;
        public int StartColumnIndex { get; init; } = OtherMethods.GetColumnIndex(startColumnLetter) - 1;
        public int EndColumnIndex { get; init; } = OtherMethods.GetColumnIndex(endColumnLetter) - 1;
    }
}