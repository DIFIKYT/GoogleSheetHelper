using Google;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System.Net;

namespace GoogleSheetsHelper
{
    public class ValuesService(SheetsService sheetsService)
    {
        private readonly SheetsService _sheetsService = sheetsService;

        /// <summary>
        /// Shell address format should be "ListName!ColumnNameRowNumber". Example: "Sheet!A1".
        /// </summary>
        /// <param name="spreadsheetId"></param>
        /// <param name="shellAdressToValue"></param>
        /// <returns></returns>
        public async Task BatchUpdateValues(string spreadsheetId, Dictionary<string, string> shellAdressToValue)
        {
            if (shellAdressToValue.Count == 0)
            {
                return;
            }

            try
            {
                const int maxBatchSize = 1000;
                const int maxRetries = 3;

                List<List<KeyValuePair<string, string>>> batches = shellAdressToValue
                    .Select((kvp, index) => new { kvp, index })
                    .GroupBy(x => x.index / maxBatchSize)
                    .Select(g => g.Select(x => x.kvp).ToList())
                    .ToList();

                foreach (List<KeyValuePair<string, string>> batch in batches)
                {
                    List<ValueRange> valueRanges = batch
                        .Select(kvp => new ValueRange { Range = kvp.Key, Values = new List<IList<object>> { new List<object> { kvp.Value } } })
                        .ToList();

                    BatchUpdateValuesRequest batchRequest = new()
                    {
                        ValueInputOption = "USER_ENTERED",
                        Data = valueRanges,
                    };

                    bool success = false;
                    int retryCount = 0;

                    while (success == false && retryCount < maxRetries)
                    {
                        try
                        {
                            SpreadsheetsResource.ValuesResource.BatchUpdateRequest request = _sheetsService.Spreadsheets.Values.BatchUpdate(batchRequest, spreadsheetId);
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
                            Console.WriteLine($"Error during batch update: {exception.Message}");
                            throw;
                        }
                    }
                }

                shellAdressToValue.Clear();
            }
            catch (Exception exception)
            {
                Console.WriteLine($"\n\nОшибка в BatchUpdateValues '{_sheetsService.Spreadsheets.Get(spreadsheetId).Execute().Properties.Title}'\n{exception}\n\n");
                shellAdressToValue.Clear();
            }
        }

        /// <summary>
        /// Range format should be "ListName!ColumnNameRowNumber:ColumnNameRowNumber". Example: "Sheet1!A1:C10".
        /// </summary>
        /// <param name="spreadsheetId"></param>
        /// <param name="range"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        public async Task BatchClearData(string spreadsheetId, List<string> clearRanges)
        {
            try
            {
                BatchClearValuesRequest requestBody = new()
                {
                    Ranges = clearRanges.ToList()
                };

                await _sheetsService!.Spreadsheets.Values.BatchClear(requestBody, spreadsheetId).ExecuteAsync();

                clearRanges.Clear();
            }
            catch (Exception exception)
            {
                Console.WriteLine("Ошибка в BatchClearData - " + exception);
            }
        }

        /// <summary>
        /// Returns cell values from the specified range in the spreadsheet.
        /// <para>
        /// Range format should be one of the following:
        /// </para>
        /// <para>
        /// • "ListName!ColumnNameRowNumber:ColumnNameRowNumber" — e.g. "Sheet1!A1:C10"
        /// </para>
        /// <para>
        /// • "ListName!ColumnNameRowNumber" — e.g. "Sheet1!A1"
        /// </para>
        /// </summary>
        /// <param name="spreadsheetId">The ID of the target spreadsheet.</param>
        /// <param name="range">The A1 notation of the cell range.</param>
        /// <returns>The requested cell values as a <see cref="ValueRange"/>.</returns>
        public ValueRange GetValueByShellsRange(string spreadsheetId, string range)
        {
            return _sheetsService.Spreadsheets.Values.Get(spreadsheetId, range).Execute();
        }

        /// <summary>
        /// Return number last row with data
        /// </summary>
        /// <param name="spreadsheetId"></param>
        /// <param name="range"></param>
        /// <returns></returns>
        public int GetLastRow(string spreadsheetId, string range)
        {
            try
            {
                ValueRange valueRange = _sheetsService.Spreadsheets.Values.Get(spreadsheetId, range).Execute();

                if (valueRange == null || valueRange.Values == null)
                    return 1;

                return valueRange.Values.Count;
            }
            catch (Exception exception)
            {
                Console.WriteLine("Ошибка получения последний строки");
                Console.WriteLine(exception);
                return -1;
            }
        }

        public int GetLastColumn(string spreadsheetId, string range)
        {
            try
            {
                ValueRange valueRange = _sheetsService.Spreadsheets.Values.Get(spreadsheetId, range).Execute();

                if (valueRange?.Values == null || valueRange.Values.Count == 0)
                    return 1;

                int maxColumns = 0;

                foreach (var row in valueRange.Values)
                {
                    if (row.Count > maxColumns)
                        maxColumns = row.Count;
                }

                return maxColumns;
            }
            catch (Exception exception)
            {
                Console.WriteLine("Ошибка получения последнего столбца");
                Console.WriteLine(exception);
                return -1;
            }
        }
        
        public Spreadsheet GetSpreadsheet(string spreadsheetId)
        {
            try
            {
                return _sheetsService!.Spreadsheets.Get(spreadsheetId).Execute();
            }
            catch (Exception exception)
            {
                Console.WriteLine("Ошибка в GetSpreadsheet - " + exception);
            }

            return new();
        }
    }
}