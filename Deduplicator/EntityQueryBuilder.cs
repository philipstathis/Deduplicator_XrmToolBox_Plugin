using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Xrm.Sdk.Query;

namespace Deduplicator
{
    internal class EntityQueryBuilder : FetchXmlBuilder
    {
        public EntityQueryBuilder()
        {
            FetchXmlHeader = new StringBuilder("<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>");
            FetchXmlAttributes = new StringBuilder();
            Attributes = new List<CrmEntityAttribute>();
            FetchXmlFilters = new StringBuilder("<filter type='and'>");
        }

        internal override void AddColumn(CrmEntityAttribute attribute)
        {
            Attributes.Add(attribute);

            // add it to the attributes String Builder
            FetchXmlAttributes.Append(
                string.Format(
                    "<attribute name='{0}' />",
                    attribute.Name));
        }

        internal override void SetEntity(string entity)
        {
            // add it to the header String Builder
            FetchXmlHeader.Append(string.Format("<entity name='{0}' >", entity));
        }

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
    }
}
