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
    class GetPlanDetails
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

                APIResponse<ResponseHandler> response = sdkOperations.GetPlanDetails();
                int responseStatusCode = response.StatusCode;

                if (responseStatusCode >= 200 && responseStatusCode <= 299)
                {
                    PlanDetails planDetails = (PlanDetails)response.Object;

                    Console.WriteLine("Plan Name - {0}", planDetails.PlanName);
                    Console.WriteLine("Subscription Period - {0}", planDetails.SubscriptionPeriod);
                    Console.WriteLine("Total Usage Limit - {0}", planDetails.UsageLimit);
                    Console.WriteLine("Total API Usage - {0}", planDetails.TotalUsage);
                    Console.WriteLine("Remaining API Limit -{0}", planDetails.RemainingUsageLimit);
                    
                }
                else
                {
                    InvalidConfigurationException invalidConfiguration = (InvalidConfigurationException)response.Object;
                    string errorMessage = invalidConfiguration.Message;

                    /*long errorCode = invalidConfiguration.Code;
                    string errorKeyName = invalidConfiguration.KeyName;
                    string errorParameterName = invalidConfiguration.ParameterName;*/

                    Console.WriteLine("Configuration error in getting plan details. Error - {0}", errorMessage);
                }
            }
            catch (System.Exception e)
            {
                Console.WriteLine("Exception in getting plan details - ", e);
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