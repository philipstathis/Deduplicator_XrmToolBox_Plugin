using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Xrm.Sdk.Query;

namespace Deduplicator
{
    class RelatedEntityQueryBuilder : FetchXmlBuilder
    {
        public RelatedEntityQueryBuilder()
        {
            FetchXmlHeader = new StringBuilder("<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>");
            FetchXmlAttributes = new StringBuilder();
            FetchXmlLinkedEntity = new StringBuilder();
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
            WorkingEntity = entity;
        }

        internal void SetRelatedEntity(string relatedEntity, string referencingAttribute)
        {
            if (string.IsNullOrWhiteSpace(WorkingEntity))
                return;

            FetchXmlHeader.Append(string.Format("<entity name='{0}' >", relatedEntity));

            AddColumn(new CrmEntityAttribute(relatedEntity + "id", "Uniqueidentifier"));
            AddColumn(new CrmEntityAttribute(referencingAttribute, "Lookup"));

            FetchXmlLinkedEntity.Append(string.Format("<link-entity name='{0}' from='{0}id' to='{1}' alias='ah' >",
                WorkingEntity, referencingAttribute));
        }

        internal override FetchExpression GetOutput()
        {
            // close Filters clause
            FetchXmlFilters.Append("</filter>");

            // put filters on the linker entity and close Link Entity
            FetchXmlLinkedEntity.Append(FetchXmlFilters);
            FetchXmlLinkedEntity.Append("</link-entity>");

            // create footer
            var fetchXmlFooter = new StringBuilder("</entity>");
            fetchXmlFooter.Append("</fetch>");

            var newquery = FetchXmlHeader.ToString() + FetchXmlAttributes + FetchXmlLinkedEntity + fetchXmlFooter;

            return new FetchExpression(newquery);
        }
    }
}
