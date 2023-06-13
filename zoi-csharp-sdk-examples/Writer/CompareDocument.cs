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
    class CompareDocument
    {
        static void execute(String[] args)
        {
            try
            {
                // Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
                // You can place SDK initializer code in your application and call once while your application start-up.
                initializeSdk();

                V1Operations sdkOperations = new V1Operations();
                CompareDocumentParameters compareParameters = new CompareDocumentParameters();

                compareParameters.Url1 = "https://demo.office-integrator.com/zdocs/MS_Word_Document_v0.docx";
                compareParameters.Url2 = "https://demo.office-integrator.com/zdocs/MS_Word_Document_v1.docx";

                String file1Name = "MS_Word_Document_v0.docx";
                String file2Name = "MS_Word_Document_v1.docx";

                /* String inputFile1Path = Path.Combine(Environment.CurrentDirectory, "sample_documents", "MS_Word_Document_v0.docx");
                StreamWrapper file1StreamWrapper = new StreamWrapper(inputFile1Path);

                compareParameters.Document1 = file1StreamWrapper;

                String inputFile2Path = Path.Combine(Environment.CurrentDirectory, "sample_documents", "MS_Word_Document_v1.docx");
                StreamWrapper file2StreamWrapper = new StreamWrapper(inputFile2Path);

                compareParameters.Document2 = file2StreamWrapper; */

                compareParameters.Lang = "en";
                compareParameters.Title = file1Name + " vs " + file2Name;

                APIResponse<WriterResponseHandler> response = sdkOperations.CompareDocument(compareParameters);
                int responseStatusCode = response.StatusCode;

                if (responseStatusCode >= 200 && responseStatusCode <= 299)
                {
                    CompareDocumentResponse compareResponse = (CompareDocumentResponse)response.Object;

                    Console.WriteLine("Compared URL - {0}", compareResponse.CompareUrl);
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
                Console.WriteLine("Exception in creating document compare url - ", e);
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