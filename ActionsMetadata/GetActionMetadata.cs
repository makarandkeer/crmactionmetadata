// <copyright file="GetActionMetadata.cs" company="">
// Copyright (c) 2014 All Rights Reserved
// </copyright>
// <author></author>
// <date>7/18/2014 1:30:37 PM</date>
// <summary>Implements the GetActionMetadata Workflow Activity.</summary>
namespace ActionsMetadata.Workflow
{
    using System;
    using System.Activities;
    using System.ServiceModel;
    using System.Xml.Linq;
    using System.Linq;

    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Query;
    using Microsoft.Xrm.Sdk.Client;
    using Microsoft.Xrm.Sdk.Workflow;

    public sealed class GetActionMetadata : CodeActivity
    {
        /// <summary>
        /// Executes the workflow activity.
        /// </summary>
        /// <param name="executionContext">The execution context.</param>
        protected override void Execute(CodeActivityContext executionContext)
        {
            // Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            if (tracingService == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve tracing service.");
            }

            tracingService.Trace("Entered GetActionMetadata.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}",
                executionContext.ActivityInstanceId,
                executionContext.WorkflowInstanceId);

            // Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();

            if (context == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve workflow context.");
            }

            tracingService.Trace("GetActionMetadata.Execute(), Correlation Id: {0}, Initiating User: {1}",
                context.CorrelationId,
                context.InitiatingUserId);

            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
               XDocument xd = ReadActionMetadata(service);

               ActionMetadataTable.Set(executionContext, xd.ToString());
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                tracingService.Trace("Exception: {0}", e.ToString());

                // Handle the exception.
                throw;
            }

            tracingService.Trace("Exiting GetActionMetadata.Execute(), Correlation Id: {0}", context.CorrelationId);
        }

        //private XDocument ReadActionMetadata(IOrganizationService _orgSvc)
        //{
        //    OrganizationServiceContext orgContext = new OrganizationServiceContext(_orgSvc);

        //    var query = from wf in orgContext.CreateQuery("workflow")
        //                join sdk in orgContext.CreateQuery("sdkmessage") on wf.GetAttributeValue<EntityReference>("sdkmessageid").Id equals sdk.GetAttributeValue<Guid>("sdkmessageid")
        //                where wf.GetAttributeValue<OptionSetValue>("statecode").Value == 1
        //                && wf.GetAttributeValue<OptionSetValue>("statuscode").Value == 2
        //                && wf.GetAttributeValue<OptionSetValue>("type").Value == 1
        //                && wf.GetAttributeValue<OptionSetValue>("category").Value == 3
        //                orderby wf.GetAttributeValue<string>("name")
        //                select new { Actions = wf, SDK = sdk };

        //    string[] htmlHeader = new string[] { "Action", "Unique Name", "Parent Entity", "Argument", "Type", "Is Mandatory", "Direction", "Is Target", "EntityAttribute", "Description" };

        //    XDocument xTable = XDocument.Parse("<table></table>");
        //    xTable.Root.Add(new XElement("tr"));
        //    foreach (var str in htmlHeader)
        //    {
        //        xTable.Root.Element("tr").Add(new XElement("td", str));
        //    }
        //    foreach (var v in query)
        //    {
        //        string xaml = v.Actions.GetAttributeValue<string>("xaml");

        //        var xHtml = ProcessXaml(xaml, v.Actions.GetAttributeValue<string>("name"),
        //            v.SDK.GetAttributeValue<string>("name"),
        //            v.Actions.GetAttributeValue<string>("primaryentity"));
        //        xTable.Root.Add(xHtml.Element("action").Elements());
        //    }

        //    return xTable;
        //}

        private XDocument ReadActionMetadata(IOrganizationService _orgSvc)
        {
            OrganizationServiceContext orgContext = new OrganizationServiceContext(_orgSvc);

            var query = from wf in orgContext.CreateQuery("workflow")
                        join sdk in orgContext.CreateQuery("sdkmessage") on wf.GetAttributeValue<EntityReference>("sdkmessageid").Id equals sdk.GetAttributeValue<Guid>("sdkmessageid")
                        where wf.GetAttributeValue<OptionSetValue>("statecode").Value == 1
                        && wf.GetAttributeValue<OptionSetValue>("statuscode").Value == 2
                        && wf.GetAttributeValue<OptionSetValue>("type").Value == 1
                        && wf.GetAttributeValue<OptionSetValue>("category").Value == 3
                        orderby wf.GetAttributeValue<string>("name")
                        select new { Actions = wf, SDK = sdk };

            string[] htmlHeader = new string[] { "Action", "Unique Name", "Parent Entity", "Argument", "Type", "Is Mandatory", "Direction", "Is Target", "EntityAttribute", "Description" };

            XDocument xTable = XDocument.Parse("<table></table>");
            xTable.Root.Add(new XElement("tr"));
            foreach (var str in htmlHeader)
            {
                xTable.Root.Element("tr").Add(new XElement("td", str));
            }
            foreach (var v in query)
            {
                string xaml = v.Actions.GetAttributeValue<string>("xaml");

                var xHtml = ProcessXaml(xaml, v.Actions.GetAttributeValue<string>("name"),
                    v.SDK.GetAttributeValue<string>("name"),
                    v.Actions.GetAttributeValue<string>("primaryentity"));
                xTable.Root.Add(xHtml.Element("action").Elements());
            }

            return xTable;
        }

        private XDocument ProcessXaml(string xaml, params string[] actionProperties)
        {
            XDocument xDom = XDocument.Parse(xaml);

            var mxswAttribute = xDom.Root.Attributes().Where(a => a.Name.LocalName == "mxsw").FirstOrDefault();

            mxswAttribute.Value = "mxsw";

            XNamespace ns = "http://schemas.microsoft.com/winfx/2006/xaml";

            XNamespace mxsw = "mxsw";

            var memElement = xDom.Descendants(ns + "Property");

            XDocument xHtml = XDocument.Parse("<action></action>");
            XDocument actionProp = XDocument.Parse("<ap></ap>");

            int propCount = memElement.Where(a => a.HasElements).Count<XElement>();

            foreach (string ap in actionProperties)
            {
                actionProp.Root.Add(
                                new XElement("td",
                                                new XAttribute("rowSpan", propCount.ToString())
                                             ) { Value = ap }
                              );
            }

            bool parentRowAdded = false;
            foreach (var prop in memElement)
            {
                if (prop.HasElements)
                {
                    string propName = prop.Attribute("Name").Value;
                    string type = prop.Attribute("Type").Value;

                    //string ArgumentRequiredAttribute = prop.Element(ns + "Property.Attributes").Element(mxsw + "ArgumentRequiredAttribute").Attribute("Value").Value;
                    //string ArgumentTargetAttribute = prop.Element(ns + "Property.Attributes").Element(mxsw + "ArgumentTargetAttribute").Attribute("Value").Value;
                    //string ArgumentDescriptionAttribute = prop.Element(ns + "Property.Attributes").Element(mxsw + "ArgumentDescriptionAttribute").Attribute("Value").Value;
                    //string ArgumentDirectionAttribute = prop.Element(ns + "Property.Attributes").Element(mxsw + "ArgumentDirectionAttribute").Attribute("Value").Value;
                    //string ArgumentEntityAttribute = prop.Element(ns + "Property.Attributes").Element(mxsw + "ArgumentEntityAttribute").Attribute("Value").Value;

                    string ArgumentRequiredAttribute = string.Empty;
                    string ArgumentTargetAttribute = string.Empty;
                    string ArgumentDescriptionAttribute = string.Empty;
                    string ArgumentDirectionAttribute = string.Empty;
                    string ArgumentEntityAttribute = string.Empty;

                    //                    List<KeyValuePair<string, string>> aList = new List<KeyValuePair<string, string>>();
                    foreach (var v in prop.Element(ns + "Property.Attributes").Elements())
                    {
                        //aList.Add(new KeyValuePair<string, string>(v.Name.LocalName, v.FirstAttribute.Value));
                        switch (v.Name.LocalName)
                        {
                            case "ArgumentRequiredAttribute":
                                ArgumentRequiredAttribute = v.FirstAttribute.Value;
                                break;
                            case "ArgumentTargetAttribute":
                                ArgumentTargetAttribute = v.FirstAttribute.Value;
                                break;
                            case "ArgumentDescriptionAttribute":
                                ArgumentDescriptionAttribute = v.FirstAttribute.Value;
                                break;
                            case "ArgumentDirectionAttribute":
                                ArgumentDirectionAttribute = v.FirstAttribute.Value;
                                break;
                            case "ArgumentEntityAttribute":
                                ArgumentEntityAttribute = v.FirstAttribute.Value;
                                break;
                        }
                    }

                    if (!parentRowAdded && actionProperties.Count() > 0)
                    {
                        xHtml.Root.Add(new XElement("tr", actionProp.Element("ap").Descendants(),
                            new XElement("td") { Value = propName },
                            new XElement("td") { Value = type },
                            new XElement("td") { Value = ArgumentRequiredAttribute },
                            new XElement("td") { Value = ArgumentDirectionAttribute },
                            new XElement("td") { Value = ArgumentTargetAttribute },
                            new XElement("td") { Value = ArgumentEntityAttribute },
                            new XElement("td") { Value = ArgumentDescriptionAttribute }
                            )
                            );
                        parentRowAdded = true;
                    }
                    else
                    {
                        xHtml.Root.Add(new XElement("tr",
                            new XElement("td") { Value = propName },
                            new XElement("td") { Value = type },
                            new XElement("td") { Value = ArgumentRequiredAttribute },
                            new XElement("td") { Value = ArgumentDirectionAttribute },
                            new XElement("td") { Value = ArgumentTargetAttribute },
                            new XElement("td") { Value = ArgumentEntityAttribute },
                            new XElement("td") { Value = ArgumentDescriptionAttribute }
                            )
                            );
                    }
                }
            }

            return xHtml;
        }


        [Output("ActionMetadataHTMLTable")]
        public OutArgument<string> ActionMetadataTable { get; set; }
    }
}