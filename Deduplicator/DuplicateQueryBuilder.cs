using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk.Query;

namespace Deduplicator
{
    internal class DuplicateQueryBuilder : FetchXmlBuilder
    {
        public DuplicateQueryBuilder()
        {
            FetchXmlHeader = new StringBuilder("<fetch mapping='logical' aggregate='true' >");
            FetchXmlAttributes = new StringBuilder();
            Attributes = new List<CrmEntityAttribute>();
            FetchXmlFilters = new StringBuilder("<filter type='and'>");
        }

        public string DuplicateColumnName { get; private set; }

        internal override FetchExpression GetOutput()
        {
            // close Filters clause
            FetchXmlFilters.Append("</filter>");

            // create footer
            var fetchXmlFooter = new StringBuilder("</entity>");
            fetchXmlFooter.Append("</fetch>");

            var newquery = FetchXmlHeader.ToString() + FetchXmlAttributes + FetchXmlFilters + fetchXmlFooter;

            return new FetchExpression(newquery);
        }

        internal void AddCustomFilterXml(string filterXml)
        {
            FetchXmlFilters.Append(filterXml);
        }

        internal override void AddColumn(CrmEntityAttribute attribute)
        {
            Attributes.Add(attribute);

            // add it to the attributes String Builder

            if (attribute.IsDate)
            {
                // DateTime requires a break down of day/month/year
                FetchXmlAttributes.Append(
                    string.Format(
                        "<attribute groupby='true' alias='GroupByDay' name='{0}' dategrouping='day' />",
                        attribute.Name));
                FetchXmlAttributes.Append(
                    string.Format(
                        "<attribute groupby='true' alias='GroupByMonth' name='{0}' dategrouping='month' />",
                        attribute.Name));
                FetchXmlAttributes.Append(
                    string.Format(
                        "<attribute groupby='true' alias='GroupByYear' name='{0}' dategrouping='year' />",
                        attribute.Name));
            }
            else
            {
                FetchXmlAttributes.Append(
                    string.Format(
                        "<attribute groupby='true' alias='GroupBy{0}' name='{0}' />",
                        attribute.Name));
            }
            FetchXmlFilters.Append(
                string.Format(
                    "<condition attribute='{0}' operator='not-null' />",
                    attribute.Name));
        }

        internal override void SetEntity(string entity)
        {
            DuplicateColumnName = "CountOf" + entity;

            // add it to the header String Builder
            FetchXmlHeader.Append(string.Format("<entity name='{0}' >", entity));

            // add the entity count clause
            FetchXmlAttributes.Append(
                string.Format(
                    "<attribute alias='CountOf{0}' name='{0}id' aggregate='count' distinct='true' />",
                    entity));
        }
    }
}
