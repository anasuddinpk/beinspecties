using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace BE_Inspecties.Workorder_Service
{
    public class CreateQuoteProduct : IPlugin
    {
        IOrganizationService service;
        ITracingService tracingService;
        IPluginExecutionContext context;
        Entity entity;
        Guid userId = Guid.Empty;

        public void Execute(IServiceProvider serviceProvider)
        {
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = serviceFactory.CreateOrganizationService(context.UserId);

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                entity = (Entity)context.InputParameters["Target"];
                Core Core = new Core(service, entity, tracingService);

                try
                {
                    if (entity.LogicalName == "msdyn_workorderservice")
                    {
                        switch (context.Stage)
                        {
                            case 10://Pre-Validation

                                switch (context.MessageName)
                                {
                                    case "Create":
                                        break;

                                    case "Update":
                                        break;

                                    default:
                                        break;
                                }
                                break;

                            case 20://Pre-Operation

                                switch (context.MessageName)
                                {
                                    case "Create":

                                        Core.CreateQuoteProduct();

                                        break;

                                    case "Update":
                                        break;

                                    default:
                                        break;
                                }
                                break;

                            case 40://Post-Operation

                                switch (context.MessageName)
                                {
                                    case "Create":
                                        break;

                                    case "Update":
                                        break;

                                    default:
                                        break;
                                }
                                break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }
    }
}
