using Com.Zoho.API.Authenticator;
using Com.Zoho.Officeintegrator;
using Com.Zoho.Officeintegrator.Dc;
using Com.Zoho.Officeintegrator.Logger;
using Com.Zoho.Officeintegrator.Util;
using Com.Zoho.Officeintegrator.V1;
using static Com.Zoho.Officeintegrator.Logger.Logger;


namespace Documents
{
    class GetDocumentInfo
    {
        static void execute(String[] args)
        {
            try
            {
                // Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
                // You can place SDK initializer code in your application and call once while your application start-up.
                initializeSdk();

                V1Operations sdkOperations = new V1Operations();
                CreateDocumentParameters parameter = new CreateDocumentParameters();

                APIResponse<WriterResponseHandler> response = sdkOperations.CreateDocument(parameter);
                int responseStatusCode = response.StatusCode;

                if (responseStatusCode >= 200 && responseStatusCode <= 299)
                {
                    CreateDocumentResponse createDocumentResponse = (CreateDocumentResponse)response.Object;
                    string documentId = createDocumentResponse.DocumentId;

                    Console.WriteLine("Document ID - {0}", documentId);

                    APIResponse<WriterResponseHandler> response1 = sdkOperations.GetDocumentInfo(documentId);

                    DocumentMeta documentMeta = (DocumentMeta)response1.Object;

                    Console.WriteLine("Document ID - {0}", documentMeta.DocumentId); //No I18N
                    Console.WriteLine("Document Name - {0}", documentMeta.DocumentName); //No I18N
                    Console.WriteLine("Document Type - {0}", documentMeta.DocumentType); //No I18N
                    Console.WriteLine("Document Expires on - {0}", documentMeta.ExpiresOn); //No I18N
                    Console.WriteLine("Document Created on - {0}", documentMeta.CreatedTime); //No I18N
                    Console.WriteLine("Active sessions count - {0}", documentMeta.CollaboratorsCount); //No I18N
                    Console.WriteLine("Collaborators count - {0}", documentMeta.CollaboratorsCount); //No I18N
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
                Console.WriteLine("Exception in getting document details - ", e);
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