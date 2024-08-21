using Com.Zoho.API.Authenticator;
using Com.Zoho.Officeintegrator;
using Com.Zoho.Officeintegrator.Dc;
using Com.Zoho.Officeintegrator.Logger;
using Com.Zoho.Officeintegrator.Util;
using Com.Zoho.Officeintegrator.V1;
using static Com.Zoho.Officeintegrator.Logger.Logger;


namespace Documents
{
    class CreateDocument
    {
        static void Main(String[] args)
        {
            try
            {
                // Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
                // You can place SDK initializer code in your application and call once while your application start-up.
                initializeSdk();

                V1Operations sdkOperations = new V1Operations();
                CreateDocumentParameters createDocumentParams = new CreateDocumentParameters();

                DocumentInfo documentInfo = new DocumentInfo();

                documentInfo.DocumentName = "Untilted Document";
                // System time value used to generate unique document every time. You can replace based on your application.
                documentInfo.DocumentId = $"{DateTimeOffset.Now.ToUnixTimeMilliseconds()}";

                createDocumentParams.DocumentInfo = documentInfo;

                UserInfo userInfo = new UserInfo();

                userInfo.UserId = "1000";
                userInfo.DisplayName = "John";

                createDocumentParams.UserInfo = userInfo;

                Margin margin = new Margin();

                margin.Top = "2in";
                margin.Bottom = "2in";
                margin.Left = "2in";
                margin.Right = "2in";

                DocumentDefaults documentDefault = new DocumentDefaults();

                documentDefault.FontSize = 14;
                documentDefault.FontName = "Arial";
                documentDefault.PaperSize = "Letter";
                documentDefault.Orientation = "portrait";
                documentDefault.TrackChanges = "disabled";

                documentDefault.Margin = margin;
                createDocumentParams.DocumentDefaults = documentDefault;

                EditorSettings editorSettings = new EditorSettings();

                editorSettings.Unit = "in";
                editorSettings.Language = "en";
                editorSettings.View = "pageview";
                createDocumentParams.EditorSettings = editorSettings;

                UiOptions uiOptions = new UiOptions();

                uiOptions.ChatPanel = "show";
                uiOptions.DarkMode = "show";
                uiOptions.FileMenu = "show";
                uiOptions.SaveButton = "show";

                createDocumentParams.UiOptions = uiOptions;

                Dictionary<string, object> permissions = new Dictionary<string, object>();

                permissions.Add("collab.chat", false);
                permissions.Add("document.edit", true);
                permissions.Add("review.comment", false);
                permissions.Add("document.export", true);
                permissions.Add("document.print", false);
                permissions.Add("document.fill", false);
                permissions.Add("review.changes.resolve", false);
                permissions.Add("document.pausecollaboration", false);

                createDocumentParams.Permissions = permissions;

                Dictionary<string, object> saveUrlParams = new Dictionary<string, object>();

                saveUrlParams.Add("id", 123456789);
                saveUrlParams.Add("auth_token", "oswedf32rk");

                Dictionary<string, object> saveUrlHeaders = new Dictionary<string, object>();

                saveUrlHeaders.Add("header1", "value1");
                saveUrlHeaders.Add("header2", "value2");

                CallbackSettings callbackSettings = new CallbackSettings();

                callbackSettings.Retries = 2;
                callbackSettings.Timeout = 10000;
                callbackSettings.SaveFormat = "docx";
                callbackSettings.HttpMethodType = "post";
                callbackSettings.SaveUrlParams = saveUrlParams;
                callbackSettings.SaveUrlHeaders = saveUrlHeaders;
                callbackSettings.SaveUrl = "https://officeintegrator.zoho.com/v1/api/webhook/savecallback/601e12157a25e63fc4dfd4e6e00cc3da2a1419fef8b7e87c0e12a5c891dc72ff";

                createDocumentParams.CallbackSettings = callbackSettings;

                APIResponse<WriterResponseHandler> response = sdkOperations.CreateDocument(createDocumentParams);
                int responseStatusCode = response.StatusCode;

                if (responseStatusCode >= 200 && responseStatusCode <= 299)
                {
                    CreateDocumentResponse documentResponse = (CreateDocumentResponse)response.Object;

                    Console.WriteLine("Document id - {0}", documentResponse.DocumentId);
                    Console.WriteLine("Document session id - {0}", documentResponse.SessionId);
                    Console.WriteLine("Document session url - {0}", documentResponse.DocumentUrl);
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
                Console.WriteLine("Exception in creating document session url - ", e);
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