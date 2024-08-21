using Com.Zoho.API.Authenticator;
using Com.Zoho.Officeintegrator;
using Com.Zoho.Officeintegrator.Dc;
using Com.Zoho.Officeintegrator.Logger;
using Com.Zoho.Officeintegrator.Util;
using Com.Zoho.Officeintegrator.V1;
using static Com.Zoho.Officeintegrator.Logger.Logger;


namespace Spreadsheet
{
    class ConvertSpreadsheet
    {
        static void execute(String[] args)
        {
            try
            {
                // Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
                // You can place SDK initializer code in your application and call once while your application start-up.
                initializeSdk();

                V1Operations sdkOperations = new V1Operations();
                SheetConversionParameters conversionParameters = new SheetConversionParameters();

                // Either use url as document source or attach the document in request body use below methods
                conversionParameters.Url = "https://demo.office-integrator.com/samples/sheet/Contact_List.xlsx";

                //String inputFilePath = Path.Combine(System.Environment.CurrentDirectory, "../../../sample_documents/Contact_List.xlsx");
                //StreamWrapper documentStreamWrapper = new StreamWrapper(inputFilePath);

                //conversionParameters.Document = documentStreamWrapper;

                SheetConversionOutputOptions outputOptions = new SheetConversionOutputOptions();

                outputOptions.Format = "pdf";
                outputOptions.DocumentName = "ConvertedSheetFile.pdf";

                conversionParameters.OutputOptions = outputOptions;

                APIResponse<SheetResponseHandler> response = sdkOperations.ConvertSheet(conversionParameters);
                int responseStatusCode = response.StatusCode;

                if (responseStatusCode >= 200 && responseStatusCode <= 299)
                {
                    FileBodyWrapper fileBodyWrapper = (FileBodyWrapper)response.Object;
                    string outputFilePath = Path.Combine(System.Environment.CurrentDirectory, "../../../sample_documents/ConvertedSheetFile.pdf");
                    using (Stream inputStream = fileBodyWrapper.File.Stream)
                    using (Stream outputStream = File.OpenWrite(outputFilePath))
                    {
                        inputStream.CopyTo(outputStream);
                    }

                    Console.WriteLine($"Converted spreadsheet saved in output file path - {outputFilePath}");
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
                Console.WriteLine("Exception in converting sheet file - ", e);
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