﻿using System;
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
    class ConvertPresentation
    {
        static void execute(String[] args)
        {
            try
            {
                // Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
                // You can place SDK initializer code in your application and call once while your application start-up.
                initializeSdk();

                V1Operations sdkOperations = new V1Operations();
                ConvertPresentationParameters conversionParameters = new ConvertPresentationParameters();

                //Either use url as document source or attach the document in request body use below methods
                conversionParameters.Url = "https://demo.office-integrator.com/samples/show/Zoho_Show.pptx";

                //String inputFilePath = "<input file path>";
                //StreamWrapper fileStreamWrapper = new StreamWrapper(inputFilePath);
                //parameters.Document = fileStreamWrapper;

                conversionParameters.Format = "pdf";

                APIResponse<ShowResponseHandler> response = sdkOperations.ConvertPresentation(conversionParameters);
                int responseStatusCode = response.StatusCode;

                if (responseStatusCode >= 200 && responseStatusCode <= 299)
                {
                    FileBodyWrapper fileBodyWrapper = (FileBodyWrapper)response.Object;
                    string outputFilePath = Path.Combine(Environment.CurrentDirectory, "ConvertedPresentationFile.pdf");
                    using (Stream inputStream = fileBodyWrapper.File.Stream)
                    using (Stream outputStream = File.OpenWrite(outputFilePath))
                    {
                        inputStream.CopyTo(outputStream);
                    }

                    Console.WriteLine($"Converted document saved in output file path - {outputFilePath}");
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
                Console.WriteLine("Exception in convering presentation - ", e);
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