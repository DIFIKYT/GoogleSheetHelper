namespace GoogleSheetsHelper
{
    public class StartEndRange(int startRow, int endRow, int startColumn, int endColumn)
    {
        public int StartRow { get; init; } = startRow;
        public int EndRow { get; init; } = endRow;
        public int StartColumn { get; init; } = startColumn;
        public int EndColumn { get; init; } = endColumn;
    }
}