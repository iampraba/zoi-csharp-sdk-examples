using System;
using Com.Zoho.Util;
using Com.Zoho.Officeintegrator.V1;
using Com.Zoho;
using Com.Zoho.Dc;
using Com.Zoho.API.Authenticator;
using Com.Zoho.API.Logger;
using static Com.Zoho.API.Logger.Logger;
using System.Collections.Generic;

namespace Writer
{
    class EditSpreadsheet
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

                parameters.Url = "https://demo.office-integrator.com/samples/sheet/Contact_List.xlsx";

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
                callbackSettings.SaveUrl = "https://officeintegrator.zoho.com/v1/api/webhook/savecallback/601e12157123434d4e6e00cc3da2406df2b9a1d84a903c6cfccf92c8286";

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
                Console.WriteLine("Exception in editing sheet - ", e);
            }
        }

        static Boolean initializeSdk()
        {
            Boolean status = false;

            try
            {
                Apikey apikey = new Apikey("2ae438cf864488657cc9754a27daa480", Com.Zoho.Util.Constants.PARAMS);
                UserSignature user = new UserSignature("john@zylker.com"); //No I18N
                Logger logger = new Logger.Builder()
                                    .Level(Levels.INFO)
                                    .FilePath("./log.txt") //No I18N
                                    .Build();

                Com.Zoho.Dc.DataCenter.Environment environment = new DataCenter.Environment("", "https://api.office-integrator.com", "", "");

                new Initializer.Builder()
                    .User(user)
                    .Environment(environment)
                    .Token(apikey)
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