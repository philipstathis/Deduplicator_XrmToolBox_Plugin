using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using Microsoft.Xrm.Sdk.Query;

namespace Deduplicator
{
    abstract class FetchXmlBuilder
    {
        internal StringBuilder FetchXmlAttributes;
        internal StringBuilder FetchXmlFilters;
        internal StringBuilder FetchXmlLinkedEntity;
        internal StringBuilder FetchXmlHeader;
        internal List<CrmEntityAttribute> Attributes;
        internal string WorkingEntity;

        internal abstract FetchExpression GetOutput();

        internal abstract void AddColumn(CrmEntityAttribute attribute);

        internal abstract void SetEntity(string entity);

        internal void AddIdentifierValuePairs(Dictionary<CrmEntityAttribute, object> valuePairs)
        {
            foreach (var pair in valuePairs)
            {
                if (pair.Key.IsDate)
                {
                    FetchXmlFilters.Append(
                        string.Format(
                            "<condition attribute='{0}' operator='on' value='{1}' />",
                            pair.Key.Name, pair.Value));
                }
                else if (pair.Key.IsLookup)
                {
                    FetchXmlFilters.Append(
                        string.Format(
                            "<condition attribute='{0}' operator='eq' value='{1}' />",
                            pair.Key.Name, pair.Value));
                }
                else
                {
                    FetchXmlFilters.Append(
                        string.Format(
                            "<condition attribute='{0}' operator='eq' value='{1}' />",
                            pair.Key.Name, pair.Value));
                }
            }
        }
    }
}
