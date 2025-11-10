using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Drive.v3;

namespace GoogleSheetsHelper
{
    public class Configurator
    {
        private static readonly string[] _sheetsScopes = [SheetsService.Scope.Spreadsheets];
        private static readonly string[] _driveScopes = [DriveService.Scope.DriveReadonly];

        public static SheetsService GetSheetService(string credentialsPath, string applicationName)
        {
            GoogleCredential sheetsCredential = GoogleCredential.FromFile(credentialsPath).CreateScoped(_sheetsScopes);

            return new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = sheetsCredential,
                ApplicationName = applicationName,
            });
        }

        public static DriveService GetDriveService(string credentialsPath, string applicationName)
        {
            GoogleCredential driveCredential = GoogleCredential.FromFile(credentialsPath).CreateScoped(_driveScopes);

            return new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = driveCredential,
                ApplicationName = applicationName,
            });
        }
    }
}