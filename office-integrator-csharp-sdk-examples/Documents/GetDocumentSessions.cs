using System;
using Com.Zoho.API.Authenticator;
using Com.Zoho.Officeintegrator;
using Com.Zoho.Officeintegrator.Dc;
using Com.Zoho.Officeintegrator.Logger;
using Com.Zoho.Officeintegrator.Util;
using Com.Zoho.Officeintegrator.V1;
using static Com.Zoho.Officeintegrator.Logger.Logger;


namespace Documents
{
    class GetDocumentSessions
    {
        static void execute(String[] args)
        {
            try
            {
                // Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
                // You can place SDK initializer code in your application and call once while your application start-up.
                initializeSdk();

                V1Operations sdkOperations = new V1Operations();
                CreateDocumentParameters createDocumentParams = new CreateDocumentParameters();

                //Either use url as document source or attach the document in request body use below methods
                createDocumentParams.Url = "https://demo.office-integrator.com/zdocs/Graphic-Design-Proposal.docx";

                DocumentInfo documentInfo = new DocumentInfo();

                documentInfo.DocumentName = "Untilted Document";
                // System time value used to generate unique document every time. You can replace based on your application.
                documentInfo.DocumentId = $"{DateTimeOffset.Now.ToUnixTimeMilliseconds()}";

                createDocumentParams.DocumentInfo = documentInfo;

                UserInfo userInfo = new UserInfo();

                userInfo.UserId = "1000";
                userInfo.DisplayName = "John";

                createDocumentParams.UserInfo = userInfo;

                APIResponse<WriterResponseHandler> response = sdkOperations.CreateDocument(createDocumentParams);
                int responseStatusCode = response.StatusCode;

                if (responseStatusCode >= 200 && responseStatusCode <= 299)
                {
                    CreateDocumentResponse documentResponse = (CreateDocumentResponse)response.Object;

                    Console.WriteLine("Document id - {0}", documentResponse.DocumentId);
                    Console.WriteLine("Document session id - {0}", documentResponse.SessionId);
                    Console.WriteLine("Document session1 url - {0}", documentResponse.DocumentUrl);

                    userInfo.UserId = "1000";
                    userInfo.DisplayName = "Praba";

                    createDocumentParams.UserInfo = userInfo;

                    response = sdkOperations.CreateDocument(createDocumentParams);

                    if (responseStatusCode >= 200 && responseStatusCode <= 299)
                    {
                        documentResponse = (CreateDocumentResponse) response.Object;

                        Console.WriteLine("Document id - {0}", documentResponse.DocumentId);
                        Console.WriteLine("Document session2 id - {0}", documentResponse.SessionId);
                        Console.WriteLine("Document session2 url - {0}", documentResponse.DocumentUrl);

                        response = sdkOperations.GetAllSessions(documentInfo.DocumentId);

                        if (responseStatusCode >= 200 && responseStatusCode <= 299)
                        {
                            AllSessionsResponse allSessionsMeta = (AllSessionsResponse)response.Object;

                            Console.WriteLine("Document id - {0}", allSessionsMeta.DocumentId);
                            Console.WriteLine("Document Name - {0}", allSessionsMeta.DocumentName);
                            Console.WriteLine("Document Type - {0}", allSessionsMeta.DocumentType);
                            Console.WriteLine("Document Expires on - {0}", allSessionsMeta.ExpiresOn);
                            Console.WriteLine("Document Created on - {0}", allSessionsMeta.CreatedTime);
                            Console.WriteLine("Active sessions count - {0}", allSessionsMeta.ActiveSessionsCount);
                            Console.WriteLine("Collaborators count - {0}", allSessionsMeta.CollaboratorsCount);
                            List<SessionMeta> sessions = allSessionsMeta.Sessions;

                            foreach (SessionMeta sessionMeta in sessions)
                            {
                                Console.WriteLine("Session status- {0}", sessionMeta.Status);
                                Console.WriteLine("Session User ID - {0}", sessionMeta.UserInfo.UserId);
                                Console.WriteLine("Session User Display Name - {0}", sessionMeta.UserInfo.DisplayName);
                                Console.WriteLine("Session Expires on - {0}", sessionMeta.Info.ExpiresOn);
                            }
                        }
                    }
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
                Console.WriteLine("Exception in getting document sessions details - ", e);
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