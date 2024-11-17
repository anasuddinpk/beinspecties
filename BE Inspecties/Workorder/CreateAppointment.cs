using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace BE_Inspecties.Workorder
{
    public class CreateAppointment : IPlugin
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

                try
                {
                    if (entity.LogicalName == "msdyn_workorder")
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

                                        if (entity.Attributes.Contains("msdyn_systemstatus") && entity.Attributes["msdyn_systemstatus"] != null)
                                        {
                                            if (entity.GetAttributeValue<OptionSetValue>("msdyn_systemstatus").Value == 690970001) //Scheduled
                                            {
                                                Entity workorderEntity = service.Retrieve("msdyn_workorder", entity.Id, new ColumnSet("msdyn_functionallocation", "msdyn_workordertype", "msdyn_primaryincidenttype", "msdyn_serviceaccount", "msdyn_pricelist", "msdyn_reportedbycontact", "msdyn_workordersummary"));
                                                Core Core = new Core(service, workorderEntity, tracingService);
                                                Guid userId = context.InitiatingUserId;
                                                Core.CreateAppointment(entity.Id, userId);
                                            }
                                        }

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
