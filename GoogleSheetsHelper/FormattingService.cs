using Google;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System.Net;

namespace GoogleSheetsHelper
{
    public class FormattingService(SheetsService sheetsService)
    {
        private readonly SheetsService _sheetsService = sheetsService;

        public async Task BatchUpdateCells(string spreadsheetId, List<Request> requests)
        {
            try
            {
                BatchUpdateSpreadsheetRequest batchUpdateRequest = new()
                {
                    Requests = requests
                };

                await _sheetsService!.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadsheetId).ExecuteAsync();

                requests.Clear();
            }
            catch (Exception exception)
            {
                Console.WriteLine("Ошибка в BatchUpdateCells - " + exception);
            }
        }

        public Request GetCopyPastInCurrentTableRequest(string spreadsheetId, int? sheetId, StartEndRange copyRange, StartEndRange pastRange)
        {
            ArgumentNullException.ThrowIfNull(sheetId);
            ArgumentNullException.ThrowIfNull(copyRange);
            ArgumentNullException.ThrowIfNull(pastRange);

            try
            {
                return new Request
                {
                    CopyPaste = new CopyPasteRequest
                    {
                        Source = new GridRange
                        {
                            SheetId = sheetId,
                            StartRowIndex = copyRange.StartRowIndex,
                            EndRowIndex = copyRange.EndRowIndex,
                            StartColumnIndex = copyRange.StartColumnIndex,
                            EndColumnIndex = copyRange.EndColumnIndex
                        },
                        Destination = new GridRange
                        {
                            SheetId = sheetId,
                            StartRowIndex = pastRange.StartRowIndex,
                            EndRowIndex = pastRange.EndRowIndex,
                            StartColumnIndex = pastRange.StartColumnIndex,
                            EndColumnIndex = pastRange.EndColumnIndex
                        },
                        PasteType = "PASTE_NORMAL",
                        PasteOrientation = "NORMAL"
                    }
                };
            }
            catch (Exception exception)
            {
                Console.WriteLine($"\n\nОшибка в CopyPastDataInCurrentTable '{_sheetsService.Spreadsheets.Get(spreadsheetId).Execute().Properties.Title}'\n{exception}\n\n");
                return new();
            }
        }

        public Request GetListAddingRequest(string spreadsheetId, string listName)
        {
            try
            {
                if (string.IsNullOrEmpty(listName))
                    throw new ArgumentNullException(nameof(listName));

                return new Request
                {
                    AddSheet = new AddSheetRequest
                    {
                        Properties = new SheetProperties
                        {
                            Title = listName,
                            Index = 0,
                            GridProperties = new GridProperties
                            {
                                RowCount = 1000,
                                ColumnCount = 26
                            }
                        }
                    }
                };
            }
            catch (Exception exception)
            {
                Console.WriteLine($"\n\nОшибка в AddNewList '{_sheetsService.Spreadsheets.Get(spreadsheetId).Execute().Properties.Title}'\n{exception}\n\n");
                return new();
            }
        }
    }
}