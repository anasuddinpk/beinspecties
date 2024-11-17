using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BE_Inspecties.Workorder
{
    public class Core
    {
        private IOrganizationService _service;
        private ITracingService _tracingService;
        private Entity _entity;

        public Core(IOrganizationService service, Entity entity, ITracingService tracingService)
        {
            _service = service;
            _entity = entity;
            _tracingService = tracingService;
        }

        public void CreateAppointment(Guid workOrderId, Guid userId)
        {
            try
            {
                _tracingService.Trace("Appointment creation started");

                var resourceBookingData = GetResourceBookingData(workOrderId);
                DateTime[] resourceBookingDates = resourceBookingData.BookingDates;
                Guid resourceId = resourceBookingData.resourceId;

                if (resourceBookingDates.Length == 2)
                {
                    String subject = "";
                    String functionalLocation = "";
                    String workorderType = "";
                    String incidentType = "";
                    String serviceAccount = "";
                    String priceList = "";
                    String reportedContactPerson = "";
                    String workorderSummary = "";
                    decimal quoteAmount = GetQuoteAmout(workOrderId);

                    Entity Appointment = new Entity("appointment");

                    // Get Functional Location
                    if (_entity.Attributes.Contains("msdyn_functionallocation") && _entity.Attributes["msdyn_functionallocation"] != null)
                    {
                        functionalLocation = GetFunctionalLocation(_entity.GetAttributeValue<EntityReference>("msdyn_functionallocation").Id);
                    }

                    // Get Workorder Type
                    if (_entity.Attributes.Contains("msdyn_workordertype") && _entity.Attributes["msdyn_workordertype"] != null)
                    {
                        workorderType = GetWorkOrderType(_entity.GetAttributeValue<EntityReference>("msdyn_workordertype").Id);
                    }

                    // Get Incident Type
                    if (_entity.Attributes.Contains("msdyn_primaryincidenttype") && _entity.Attributes["msdyn_primaryincidenttype"] != null)
                    {
                        incidentType = GetIncidentType(_entity.GetAttributeValue<EntityReference>("msdyn_primaryincidenttype").Id);
                    }

                    // Get Service Account
                    if (_entity.Attributes.Contains("msdyn_serviceaccount") && _entity.Attributes["msdyn_serviceaccount"] != null)
                    {
                        serviceAccount = GetServiceAccount(_entity.GetAttributeValue<EntityReference>("msdyn_serviceaccount").Id);
                    }

                    // Get Price List
                    if (_entity.Attributes.Contains("msdyn_pricelist") && _entity.Attributes["msdyn_pricelist"] != null)
                    {
                        priceList = GetPriceList(_entity.GetAttributeValue<EntityReference>("msdyn_pricelist").Id);
                    }

                    // Get Reported Contact Person
                    if (_entity.Attributes.Contains("msdyn_reportedbycontact") && _entity.Attributes["msdyn_reportedbycontact"] != null)
                    {
                        reportedContactPerson = GetReportedContactPerson(_entity.GetAttributeValue<EntityReference>("msdyn_reportedbycontact").Id);
                    }

                    // Get Workorder Summary
                    if (_entity.Attributes.Contains("msdyn_workordersummary") && _entity.Attributes["msdyn_workordersummary"] != null)
                    {
                        workorderSummary = _entity.GetAttributeValue<string>("msdyn_workordersummary");
                    }

                    // Compose Subject of Appointment
                    subject = $"{functionalLocation} - {workorderType} - {quoteAmount.ToString("F2")}".Replace("-$", "").Replace("$-", "");

                    Appointment["subject"] = subject;
                    Appointment["scheduledstart"] = resourceBookingDates[0];
                    Appointment["scheduledend"] = resourceBookingDates[1];
                    Appointment["regardingobjectid"] = new EntityReference("msdyn_workorder", workOrderId);
                    Appointment["location"] = functionalLocation;

                    EntityCollection requiredAttendee = new EntityCollection();
                    Entity requiredAttendeeLookup = new Entity("activityparty");
                    requiredAttendeeLookup["partyid"] = new EntityReference("systemuser", userId);
                    requiredAttendee.Entities.Add(requiredAttendeeLookup);
                    Appointment["requiredattendees"] = requiredAttendee;

                    if (resourceId != Guid.Empty)
                    {
                        Guid resourceUserId = GetBookableResourceUser(resourceId);

                        if (resourceUserId != Guid.Empty)
                        {
                            EntityCollection organizer = new EntityCollection();
                            Entity oganizerLookup = new Entity("activityparty");
                            oganizerLookup["partyid"] = new EntityReference("systemuser", resourceUserId);
                            organizer.Entities.Add(oganizerLookup);
                            Appointment["organizer"] = organizer;
                        }
                    }

                    string description = "Service Account: " + serviceAccount +
                                         "\nWorkOrder Type: " + workorderType +
                                         "\nIncident Type: " + incidentType +
                                         "\nSummary: " + workorderSummary +
                                         "\nPrice List: " + priceList +
                                         "\nReported Contact Person: " + reportedContactPerson;

                    Appointment["description"] = description;

                    _service.Create(Appointment);

                    _tracingService.Trace("Appointment created successfully.");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException($"Error creating Appointment: {ex.Message}");
            }
        }

        private String GetFunctionalLocation(Guid functionalLocationId)
        {
            String functionalLocation = "";

            Entity functionalLocationEntity = _service.Retrieve("msdyn_functionallocation", functionalLocationId, new ColumnSet("msdyn_name"));

            if (functionalLocationEntity.Attributes.Contains("msdyn_name") && functionalLocationEntity.Attributes["msdyn_name"] != null)
            {
                functionalLocation = functionalLocationEntity.GetAttributeValue<string>("msdyn_name");
            }

            return functionalLocation;
        }

        private String GetWorkOrderType(Guid workorderTypeId)
        {
            String workorderType = "";

            Entity workorderTypeEntity = _service.Retrieve("msdyn_workordertype", workorderTypeId, new ColumnSet("msdyn_name"));

            if (workorderTypeEntity.Attributes.Contains("msdyn_name") && workorderTypeEntity.Attributes["msdyn_name"] != null)
            {
                workorderType = workorderTypeEntity.GetAttributeValue<string>("msdyn_name");
            }

            return workorderType;
        }

        private String GetIncidentType(Guid incidentTypeId)
        {
            String incidentType = "";

            Entity incidentTypeEntity = _service.Retrieve("msdyn_incidenttype", incidentTypeId, new ColumnSet("msdyn_name"));

            if (incidentTypeEntity.Attributes.Contains("msdyn_name") && incidentTypeEntity.Attributes["msdyn_name"] != null)
            {
                incidentType = incidentTypeEntity.GetAttributeValue<string>("msdyn_name");
            }

            return incidentType;
        }

        private String GetPriceList(Guid priceListId)
        {
            String priceList = "";

            Entity priceListEntity = _service.Retrieve("pricelevel", priceListId, new ColumnSet("name"));

            if (priceListEntity.Attributes.Contains("name") && priceListEntity.Attributes["name"] != null)
            {
                priceList = priceListEntity.GetAttributeValue<string>("name");
            }

            return priceList;
        }

        private String GetReportedContactPerson(Guid reportedContactPersonId)
        {
            String reportedContactPerson = "";

            Entity reportedContactPersonEntity = _service.Retrieve("contact", reportedContactPersonId, new ColumnSet("fullname"));

            if (reportedContactPersonEntity.Attributes.Contains("fullname") && reportedContactPersonEntity.Attributes["fullname"] != null)
            {
                reportedContactPerson = reportedContactPersonEntity.GetAttributeValue<string>("fullname");
            }

            return reportedContactPerson;
        }

        private String GetServiceAccount(Guid serviceAccountId)
        {
            String serviceAccount = "";

            Entity serviceAccountEntity = _service.Retrieve("account", serviceAccountId, new ColumnSet("name"));

            if (serviceAccountEntity.Attributes.Contains("name") && serviceAccountEntity.Attributes["name"] != null)
            {
                serviceAccount = serviceAccountEntity.GetAttributeValue<string>("name");
            }

            return serviceAccount;
        }

        private (Guid resourceId, DateTime[] BookingDates) GetResourceBookingData(Guid workOrderId)
        {
            List<DateTime> dates = new List<DateTime>();
            Guid resourceId = Guid.Empty;

            //QueryExpression queryForBookingStatus = new QueryExpression("bookingstatus");
            //queryForBookingStatus.ColumnSet = new ColumnSet("name");
            //queryForBookingStatus.Criteria.AddCondition("name", ConditionOperator.Equal, "Reservering"); //Booking Status equals to Reserving
            //EntityCollection bookingStatus = _service.RetrieveMultiple(queryForBookingStatus);

            QueryExpression query = new QueryExpression("bookableresourcebooking");
            query.ColumnSet = new ColumnSet("starttime", "endtime", "resource");
            query.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, workOrderId);
            //query.Criteria.AddCondition("bookingstatus", ConditionOperator.Equal, bookingStatus.Entities.First().Id);

            EntityCollection resourceBookings = _service.RetrieveMultiple(query);

            if (resourceBookings != null && resourceBookings.Entities.Count > 0)
            {
                if (resourceBookings.Entities.First().Attributes.Contains("starttime") && resourceBookings.Entities.First().Attributes["starttime"] != null)
                {
                    dates.Add(resourceBookings.Entities.First().GetAttributeValue<DateTime>("starttime"));
                }

                if (resourceBookings.Entities.First().Attributes.Contains("endtime") && resourceBookings.Entities.First().Attributes["endtime"] != null)
                {
                    dates.Add(resourceBookings.Entities.First().GetAttributeValue<DateTime>("endtime"));
                }

                if (resourceBookings.Entities.First().Attributes.Contains("resource") && resourceBookings.Entities.First().Attributes["resource"] != null)
                {
                    resourceId = resourceBookings.Entities.First().GetAttributeValue<EntityReference>("resource").Id;
                }
            }

            return (resourceId, dates.ToArray());
        }

        private Guid GetBookableResourceUser(Guid resourceId)
        {
            Guid resourceUserId = Guid.Empty;

            Entity resourceUser = _service.Retrieve("bookableresource", resourceId, new ColumnSet("userid"));

            if (resourceUser.Attributes.Contains("userid") && resourceUser.Attributes["userid"] != null)
            {
                resourceUserId = resourceUser.GetAttributeValue<EntityReference>("userid").Id;
            }

            return resourceUserId;
        }

        private decimal GetQuoteAmout(Guid workOrderId)
        {
            decimal quoteAmount = 0m;

            QueryExpression query = new QueryExpression("quote");
            query.ColumnSet = new ColumnSet("totalamount");

            query.Criteria.AddCondition("new_workorder", ConditionOperator.Equal, workOrderId);

            EntityCollection quotes = _service.RetrieveMultiple(query);

            if (quotes.Entities != null && quotes.Entities.Count > 0)
            {
                if (quotes.Entities.First().Attributes.Contains("totalamount") && quotes.Entities.First().Attributes["totalamount"] != null)
                {
                    quoteAmount = quotes.Entities.First().GetAttributeValue<Money>("totalamount").Value;
                }
            }

            return quoteAmount;
        }
    }
}
