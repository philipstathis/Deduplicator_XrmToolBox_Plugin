using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;

namespace Deduplicator
{
    internal class CrmReader
    {
        internal object GetEntityReferenceId(Entity response, string attributeName)
        {
            var entityReference = GetValue(response, attributeName) as EntityReference;
            if (entityReference != null)
            {
                return entityReference.Id;
            }
            return null;
        }

        internal string GetEntityReferenceName(Entity response, string attributeName)
        {
            var entityReference = GetValue(response, attributeName) as EntityReference;
            return entityReference != null ? entityReference.Name : null;
        }

        internal Guid GetEntityReferenceValue(Entity response, string attributeName)
        {
            var entityReference = GetValue(response, attributeName) as EntityReference;
            return entityReference != null ? entityReference.Id : Guid.Empty;
        }

        internal object GetValue(Entity entry, string attributeName)
        {
            return entry.Attributes.Contains(attributeName) ? entry.GetAttributeValue<object>(attributeName) : null;
        }

        internal object GetAliasedValue(Entity entry, string attributeName)
        {
            if (!entry.Attributes.Contains(attributeName)) return null;
            var aliasedValue = entry.GetAttributeValue<AliasedValue>(attributeName);
            if (aliasedValue == null) return null;
            var entityReferenceValue = aliasedValue.Value as EntityReference;
            return entityReferenceValue != null ? entityReferenceValue.Id : aliasedValue.Value;
        }


        internal object GetAliasedName(Entity entry, string attributeName)
        {
            if (!entry.Attributes.Contains(attributeName)) return null;
            var aliasedValue = entry.GetAttributeValue<AliasedValue>(attributeName);
            if (aliasedValue == null) return null;
            var entityReferenceValue = aliasedValue.Value as EntityReference;
            return entityReferenceValue != null ? entityReferenceValue.Name : aliasedValue.Value;
        }
    }
}
