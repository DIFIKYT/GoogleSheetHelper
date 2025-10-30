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
            if (requests == null || requests.Count == 0)
                return;

            try
            {
                const int maxBatchSize = 1000;
                const int maxRetries = 3;

                List<List<Request>> batches = requests
                    .Select((req, index) => new { req, index })
                    .GroupBy(x => x.index / maxBatchSize)
                    .Select(g => g.Select(x => x.req).ToList())
                    .ToList();

                foreach (List<Request> batch in batches)
                {
                    BatchUpdateSpreadsheetRequest batchRequest = new()
                    {
                        Requests = batch
                    };

                    bool success = false;
                    int retryCount = 0;

                    while (success == false && retryCount < maxRetries)
                    {
                        try
                        {
                            SpreadsheetsResource.BatchUpdateRequest request =
                                _sheetsService.Spreadsheets.BatchUpdate(batchRequest, spreadsheetId);

                            await request.ExecuteAsync();
                            success = true;

                            if (batches.Count > 1)
                                await Task.Delay(1000);
                        }
                        catch (GoogleApiException ex) when (ex.HttpStatusCode == HttpStatusCode.TooManyRequests)
                        {
                            retryCount++;

                            if (retryCount >= maxRetries)
                            {
                                Console.WriteLine($"Failed after {maxRetries} retries: {ex.Message}");
                                throw;
                            }

                            int delay = (int)Math.Pow(2, retryCount) * 1000;
                            await Task.Delay(delay);
                        }
                        catch (Exception exception)
                        {
                            Console.WriteLine($"Error during batch cell update: {exception.Message}");
                            throw;
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine($"\n\nОшибка в BatchUpdateCells '{_sheetsService.Spreadsheets.Get(spreadsheetId).Execute().Properties.Title}'\n{exception}\n\n");
            }
        }

        public Request GetCopyPastInCurrentTableRequest(string spreadsheetId, int sheetId, List<Request> requests, StartEndRange copyRange, StartEndRange pastRange)
        {
            ArgumentNullException.ThrowIfNull(sheetId);
            ArgumentNullException.ThrowIfNull(requests);
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
                            StartRowIndex = copyRange.StartRow,
                            EndRowIndex = copyRange.EndRow,
                            StartColumnIndex = copyRange.StartColumn,
                            EndColumnIndex = copyRange.EndColumn
                        },
                        Destination = new GridRange
                        {
                            SheetId = sheetId,
                            StartRowIndex = pastRange.StartRow,
                            EndRowIndex = pastRange.EndRow,
                            StartColumnIndex = pastRange.StartColumn,
                            EndColumnIndex = pastRange.EndColumn
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