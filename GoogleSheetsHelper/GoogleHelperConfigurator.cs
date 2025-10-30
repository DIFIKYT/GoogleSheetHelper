using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;

namespace GoogleSheetsHelper
{
    public class GoogleHelperConfigurator
    {
        private static readonly string[] _sheetsScopes = [SheetsService.Scope.Spreadsheets];

        public SheetsService GetSheetService(string credentialsPath, string applicationName)
        {
            GoogleCredential spreadsheetsCredential = GoogleCredential.FromFile(credentialsPath).CreateScoped(_sheetsScopes);

            return new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = spreadsheetsCredential,
                ApplicationName = applicationName,
            });
        }
    }
}