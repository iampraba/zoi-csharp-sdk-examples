using System;
using Com.Zoho.Util;
using Com.Zoho.Officeintegrator.V1;
using Com.Zoho;
using Com.Zoho.Dc;
using Com.Zoho.API.Authenticator;
using Com.Zoho.API.Logger;
using static Com.Zoho.API.Logger.Logger;
using System.Collections.Generic;

namespace Presentation
{
    class EditPresentation
    {
        static void execute(String[] args)
        {
            try
            {
                // Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
                // You can place SDK initializer code in your application and call once while your application start-up.
                initializeSdk();

                V1Operations sdkOperations = new V1Operations();
                CreatePresentationParameters parameters = new CreatePresentationParameters();

                parameters.Url = "https://demo.office-integrator.com/samples/show/Zoho_Show.pptx";

                //String inputFilePath = "<input file path>";
                //StreamWrapper fileStreamWrapper = new StreamWrapper(inputFilePath);
                //parameters.Document = fileStreamWrapper;

                DocumentInfo documentInfo = new DocumentInfo();

                documentInfo.DocumentName = "Untilted Presentation";
                // System time value used to generate unique document every time. You can replace based on your application.
                documentInfo.DocumentId = $"{DateTimeOffset.Now.ToUnixTimeMilliseconds()}";

                parameters.DocumentInfo = documentInfo;

                UserInfo userInfo = new UserInfo();

                userInfo.DisplayName = "John";
                userInfo.UserId = "100";

                parameters.UserInfo = userInfo;

                ZohoShowEditorSettings editorSettings = new ZohoShowEditorSettings();

                editorSettings.Language = "en";

                parameters.EditorSettings = editorSettings;

                Dictionary<string, object> permissions = new Dictionary<string, object>();

                permissions.Add("document.export", true);
                permissions.Add("document.print", true);

                parameters.Permissions = permissions;

                Dictionary<string, object> saveUrlParams = new Dictionary<string, object>();

                saveUrlParams.Add("id", 123456789);
                saveUrlParams.Add("auth_token", "oswedf32rk");

                ShowCallbackSettings callbackSettings = new ShowCallbackSettings();

                callbackSettings.SaveFormat = "pptx";
                callbackSettings.SaveUrl = "https://officeintegrator.zoho.com/v1/api/webhook/savecallback/601e12157123434d4e6e00cc3da2406df2b9a1d84a903c6cfccf92c8286";

                parameters.CallbackSettings = callbackSettings;

                APIResponse<ShowResponseHandler> response = sdkOperations.CreatePresentation(parameters);
                int responseStatusCode = response.StatusCode;

                if (responseStatusCode >= 200 && responseStatusCode <= 299)
                {
                    CreateDocumentResponse showResponse = (CreateDocumentResponse)response.Object;

                    Console.WriteLine("Presentation id - {0}", showResponse.DocumentId);
                    Console.WriteLine("Presentation session id - {0}", showResponse.SessionId);
                    Console.WriteLine("Presentation session url - {0}", showResponse.DocumentUrl);
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
                Console.WriteLine("Exception in opening presentation for editing - ", e);
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