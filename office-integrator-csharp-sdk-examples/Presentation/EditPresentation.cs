using Com.Zoho.API.Authenticator;
using Com.Zoho.Officeintegrator;
using Com.Zoho.Officeintegrator.Dc;
using Com.Zoho.Officeintegrator.Logger;
using Com.Zoho.Officeintegrator.Util;
using Com.Zoho.Officeintegrator.V1;
using static Com.Zoho.Officeintegrator.Logger.Logger;


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

                //Either use url as document source or attach the document in request body use below methods
                parameters.Url = "https://demo.office-integrator.com/samples/show/Zoho_Show.pptx";

                //String inputFilePath = Path.Combine(System.Environment.CurrentDirectory, "../../../sample_documents/Zoho_Show.pptx");
                //StreamWrapper documentStreamWrapper = new StreamWrapper(inputFilePath);

                //parameters.Document = documentStreamWrapper;

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

                CallbackSettings callbackSettings = new CallbackSettings();

                callbackSettings.SaveFormat = "pptx";
                callbackSettings.SaveUrlParams = saveUrlParams;
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