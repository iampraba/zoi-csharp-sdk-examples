using Com.Zoho.API.Authenticator;
using Com.Zoho.Officeintegrator;
using Com.Zoho.Officeintegrator.Dc;
using Com.Zoho.Officeintegrator.Logger;
using Com.Zoho.Officeintegrator.Util;
using Com.Zoho.Officeintegrator.V1;
using static Com.Zoho.Officeintegrator.Logger.Logger;

namespace Documents
{
    class MergeAndDeliver
    {
        static void execute(String[] args)
        {
            try
            {
                // Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
                // You can place SDK initializer code in your application and call once while your application start-up.
                initializeSdk();

                V1Operations sdkOperations = new V1Operations();
                MergeAndDeliverViaWebhookParameters parameters = new MergeAndDeliverViaWebhookParameters();

                //Either use url as document source or attach the document in request body use below methods
                parameters.FileUrl = "https://demo.office-integrator.com/zdocs/OfferLetter.zdoc";
                parameters.MergeDataJsonUrl = "https://demo.office-integrator.com/data/candidates.json";

                //String inputFilePath = Path.Combine(System.Environment.CurrentDirectory, "../../../sample_documents/OfferLetter.zdoc");
                //StreamWrapper documentStreamWrapper = new StreamWrapper(inputFilePath);

                //parameters.FileContent = documentStreamWrapper;

                //String mergeDataJsonFilePath = Path.Combine(System.Environment.CurrentDirectory, "../../../sample_documents/candidates.json");
                //StreamWrapper mergeDataJsonFileStreamWrapper = new StreamWrapper(mergeDataJsonFilePath);

                //parameters.MergeDataJsonContent = mergeDataJsonFileStreamWrapper;

                DocumentConversionOutputOptions outputOptions = new DocumentConversionOutputOptions();

                parameters.OutputFormat = "zdoc";
                parameters.MergeTo = "separatedoc";
                parameters.Password = "***";

                MailMergeWebhookSettings webhookSettings = new MailMergeWebhookSettings();

                webhookSettings.InvokeUrl = "https://officeintegrator.zoho.com/v1/api/webhook/savecallback/601e12157a25e63fc4dfd4e6e00cc3da2406df2b9a1d84a903c6cfccf92c8286";
                webhookSettings.InvokePeriod = "oncomplete";

                parameters.Webhook = webhookSettings;

                APIResponse<WriterResponseHandler> response = sdkOperations.MergeAndDeliverViaWebhook(parameters);
                int responseStatusCode = response.StatusCode;

                if (responseStatusCode >= 200 && responseStatusCode <= 299)
                {
                    MergeAndDeliverViaWebhookSuccessResponse mergeResponse = (MergeAndDeliverViaWebhookSuccessResponse)response.Object;

                    Console.WriteLine("Total Records Count - {0}", mergeResponse.Records.Count);
                    Console.WriteLine("Total Report URL - {0}",  mergeResponse.MergeReportDataUrl);
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
                Console.WriteLine("Exception in merging document - ", e);
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