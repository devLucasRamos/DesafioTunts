using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using Google.Apis.Auth.OAuth2;

namespace GoogleAPI
{
    class SheetsServices
    {
        private const string applicationName = "APIGoogle";

        public static SheetsService GetSheetsService(UserCredential credential)
        {

            var serviceInitializer = new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName
            };

            return new SheetsService(serviceInitializer);
        }

    }
}
