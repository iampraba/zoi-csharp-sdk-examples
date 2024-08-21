using Com.Zoho.API.Authenticator;
using Com.Zoho.Officeintegrator;
using Com.Zoho.Officeintegrator.Dc;
using Com.Zoho.Officeintegrator.Logger;
using Com.Zoho.Officeintegrator.Util;
using Com.Zoho.Officeintegrator.V1;
using static Com.Zoho.Officeintegrator.Logger.Logger;


namespace Spreadsheet
{
    class CreateSpreadsheet
    {
        static void execute(String[] args)
        {
            try
            {
                // Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
                // You can place SDK initializer code in your application and call once while your application start-up.
                initializeSdk();

                V1Operations sdkOperations = new V1Operations();
                CreateSheetParameters parameters = new CreateSheetParameters();

                DocumentInfo documentInfo = new DocumentInfo();

                documentInfo.DocumentName = "Untilted Spreadsheet";
                // System time value used to generate unique document every time. You can replace based on your application.
                documentInfo.DocumentId = $"{DateTimeOffset.Now.ToUnixTimeMilliseconds()}";

                parameters.DocumentInfo = documentInfo;

                SheetUserSettings userSettings = new SheetUserSettings();

                userSettings.DisplayName = "John";

                parameters.UserInfo = userSettings;

                DocumentDefaults documentDefault = new DocumentDefaults();

                SheetEditorSettings editorSettings = new SheetEditorSettings();

                editorSettings.Language = "en";
                editorSettings.Country = "US";

                parameters.EditorSettings= editorSettings;

                SheetUiOptions uiOptions = new SheetUiOptions();

                uiOptions.SaveButton = "show";

                parameters.UiOptions = uiOptions;

                Dictionary<string, object> permissions = new Dictionary<string, object>();

                permissions.Add("document.export", true);
                permissions.Add("document.print", true);

                parameters.Permissions = permissions;

                Dictionary<string, object> saveUrlParams = new Dictionary<string, object>();

                saveUrlParams.Add("id", 123456789);
                saveUrlParams.Add("auth_token", "oswedf32rk");

                Dictionary<string, object> saveUrlHeaders = new Dictionary<string, object>();

                saveUrlHeaders.Add("header1", "value1");
                saveUrlHeaders.Add("header2", "value2");

                SheetCallbackSettings callbackSettings = new SheetCallbackSettings();

                callbackSettings.SaveFormat = "xlsx";
                callbackSettings.SaveUrlParams = saveUrlParams;
                callbackSettings.SaveUrlHeaders = saveUrlHeaders;
                callbackSettings.SaveUrl = "https://122a4a0a4b36d2e30488e6700fbb3ca6.m.pipedream.net";

                parameters.CallbackSettings = callbackSettings;

                APIResponse<SheetResponseHandler> response = sdkOperations.CreateSheet(parameters);
                int responseStatusCode = response.StatusCode;

                if (responseStatusCode >= 200 && responseStatusCode <= 299)
                {
                    CreateSheetResponse sheetResponse = (CreateSheetResponse)response.Object;

                    Console.WriteLine("Sheet id - {0}", sheetResponse.DocumentId);
                    Console.WriteLine("Sheet session id - {0}", sheetResponse.SessionId);
                    Console.WriteLine("Sheet session url - {0}", sheetResponse.DocumentUrl);
                }
                else
                {
                    InvalidConfigurationException invalidConfiguration = (InvalidConfigurationException)response.Object;
                    string errorMessage = invalidConfiguration.Message;

                    /*long errorCode = invalidConfiguration.Code;
                    string errorKeyName = invalidConfiguration.KeyName;
                    string errorParameterName = invalidConfiguration.ParameterName;*/

                    Console.WriteLine("configuration error - {0}", errorMessage);
                }
            }
            catch (System.Exception e)
            {
                Console.WriteLine("Exception in creating sheet session url - ", e);
            }
        }

        static Boolean initializeSdk()
        {
            Boolean status = false;

            try
            {
                //Sdk application log configuration
                Logger logger = new Logger.Builder()
                        .Level(Levels.INFO)
                        //.filePath("<file absolute path where logs would be written>") //No I18N
                        .Build();

                List<IToken> tokens = new List<IToken>();
                Auth auth = new Auth.Builder()
                    .AddParam("apikey", "2ae438cf864488657cc9754a27daa480") //Update this apikey with your own apikey signed up in office inetgrator service
                    .AuthenticationSchema(new Authentication.TokenFlow())
                    .Build();

                tokens.Add(auth);

                Com.Zoho.Officeintegrator.Dc.Environment environment = new APIServer.Production("https://api.office-integrator.com"); // Refer this help page for api end point domain details -  https://www.zoho.com/officeintegrator/api/v1/getting-started.html

                new Initializer.Builder()
                    .Environment(environment)
                    .Tokens(tokens)
                    .Logger(logger)
                    .Initialize();

                status = true;
            }
            catch (System.Exception e)
            {
                Console.WriteLine("Exception in Init SDK", e);
            }
            return status;
        }
    }
}