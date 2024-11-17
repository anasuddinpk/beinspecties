using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IdentityModel.Metadata;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace BE_Inspecties.Workorder_Service
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

        //Create Quote Product on Workorder Service creation
        public void CreateQuoteProduct()
        {
            try
            {
                Guid workOrderId = GetWorkOrder(_entity);
                Guid productId = Guid.Empty;
                string productName = "";

                if (workOrderId != Guid.Empty)
                {
                    Entity product = GetProduct();

                    if (product != null)
                    {
                        productId = product.Id;
                        productName = product.GetAttributeValue<string>("name");
                    }

                    Entity QuoteProduct = new Entity("quotedetail");

                    QuoteProduct["quotedetailname"] = productName;
                    QuoteProduct["quantity"] = 1.0m;
                    QuoteProduct["productid"] = new EntityReference("product", productId);
                    QuoteProduct["uomid"] = new EntityReference("uom", new Guid("8a33926b-3727-4dbf-8ae2-d9ef6de5060c"));
                    QuoteProduct["producttypecode"] = new OptionSetValue(1);
                    QuoteProduct["quoteid"] = new EntityReference("quote", GetQuote(workOrderId));

                    if (_entity.Attributes.Contains("msdyn_pricelist") && _entity.Attributes["msdyn_pricelist"] != null)
                    {
                        QuoteProduct["msdyn_pricelist"] = new EntityReference("pricelevel", _entity.GetAttributeValue<EntityReference>("msdyn_pricelist").Id);
                    }

                    Guid quoteProductId = _service.Create(QuoteProduct);

                    _tracingService.Trace("Quote Product record created successfully, Id: " + quoteProductId);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException($"Error creating Quote Product: {ex.Message}");
            }
        }

        private Guid GetWorkOrder(Entity workOrderService)
        {
            if (workOrderService.Attributes.Contains("msdyn_workorder") && workOrderService.Attributes["msdyn_workorder"] != null)
            {
                EntityReference workOrder = workOrderService.GetAttributeValue<EntityReference>("msdyn_workorder");
                return workOrder.Id;
            }
            else
            {
                return Guid.Empty;
            }
        }

        private Entity GetProduct()
        {
            Entity product = null;

            if (_entity.Attributes.Contains("msdyn_service") && _entity.Attributes["msdyn_service"] != null)
            {
                Guid productId = _entity.GetAttributeValue<EntityReference>("msdyn_service").Id;

                product = _service.Retrieve("product", productId, new ColumnSet("name"));
            }

            return product;
        }

        private Guid GetQuote(Guid workOrderId)
        {
            Guid quoteId = Guid.Empty;

            QueryExpression query = new QueryExpression("quote");
            query.ColumnSet = new ColumnSet("quoteid");

            query.Criteria.AddCondition("new_workorder", ConditionOperator.Equal, workOrderId);

            EntityCollection quotes = _service.RetrieveMultiple(query);

            if (quotes.Entities != null && quotes.Entities.Count > 0)
            {
                quoteId = quotes.Entities.First().Id;
            }

            return quoteId;
        }

        //Delete Quote Product on Workorder Service delete
        public void DeleteQuoteProduct(Entity WorkOrderService_PreImageEntity)
        {
            try
            {
                Guid workOrderId = GetWorkOrder(WorkOrderService_PreImageEntity);

                if (workOrderId != Guid.Empty)
                {
                    Guid quoteId = GetQuote(workOrderId);
                    Guid productId = GetProduct().Id;

                    if (quoteId != null && productId != Guid.Empty)
                    {
                        QueryExpression query = new QueryExpression("quotedetail");
                        query.ColumnSet = new ColumnSet("quotedetailid");

                        query.Criteria.AddCondition("quoteid", ConditionOperator.Equal, quoteId);
                        query.Criteria.AddCondition("productid", ConditionOperator.Equal, productId);

                        EntityCollection quoteProducts = _service.RetrieveMultiple(query);

                        if (quoteProducts.Entities != null && quoteProducts.Entities.Count > 0)
                        {
                            _service.Delete("quotedetail", quoteProducts.Entities.First().Id);
                            _tracingService.Trace("Quote Product record deleted successfully, Id: " + quoteProducts.Entities.First().Id);
                        }
                        else
                        {
                            throw new InvalidPluginExecutionException("No Quote Product found with the specified quoteid and productid.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException($"Error deleting Quote Product: {ex.Message}");
            }
        }
    }
}
