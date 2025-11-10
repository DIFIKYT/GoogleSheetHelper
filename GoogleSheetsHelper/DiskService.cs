using Google.Apis.Drive.v3;

namespace GoogleSheetsHelper
{
    public class DiskService(DriveService driveService)
    {
        private readonly DriveService _driveService = driveService;

        public List<string> GetSpreadshetIds(string googleFolderId, string spreadsheetFilterWord)
        {
            try
            {
                var request = _driveService!.Files.List();
                request.Q = $"'{googleFolderId}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false";
                request.Fields = "files(id, name)";

                var result = request.Execute();

                return result.Files.Where(item => item.Name.StartsWith(spreadsheetFilterWord)).Select(item => item.Id).ToList<string>();
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Exepction in {nameof(GetSpreadshetIds)}:\n{exception}");
                return [];
            }
        }
    }
}