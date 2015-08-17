using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace Deduplicator
{
    class EntityMetadataReader : CrmReader
    {
        private readonly RetrieveAllEntitiesResponse _metadataResponse;
        private OneToManyRelationshipMetadata[] _entityRelationshipMetadataAttributeList;

        public EntityMetadataReader(OrganizationResponse response)
        {
            var metadataResponse = response as RetrieveAllEntitiesResponse;

            if (metadataResponse == null)
                throw new ArgumentException("NULL metadata response, no valid response from CRM");

            _metadataResponse = metadataResponse;
        }
        internal List<string> ParseEntityList()
        {
            var output = _metadataResponse.EntityMetadata.Select(entityMetadata => entityMetadata.LogicalName)
                .ToList();

            output.Sort();

            return output;
        }

        internal object ParseAttributes(string entitySelected)
        {
            var sampleDataSet = new DataSet {Locale = CultureInfo.InvariantCulture};
            DataTable output = sampleDataSet.Tables.Add("Attributes");

            output.Columns.Add("Logical Name", typeof (string));
            output.Columns.Add("Attribute Type", typeof(string));
            output.Columns.Add("Set As Unique Identifier?", typeof (bool));
            output.Columns.Add("Display in Step 5?", typeof (bool));

            var attributeList = _metadataResponse.EntityMetadata
                .Where(
                    entityMetadata =>
                        string.Equals(entitySelected, entityMetadata.LogicalName, StringComparison.OrdinalIgnoreCase))
                .Select(entityMetadata => entityMetadata.Attributes)
                .First();
            foreach (var attribute in attributeList)
            {
                var row = output.NewRow();

                row["Logical Name"] = attribute.LogicalName;
                row["Attribute Type"] = attribute.AttributeType;
                row["Set As Unique Identifier?"] = false;
                row["Display in Step 5?"] = false;
                output.Rows.Add(row);
            }

            output.AcceptChanges();

            // add sorting
            return output;
        }

        internal object ParseRelationships(string entitySelected)
        {
            _entityRelationshipMetadataAttributeList = _metadataResponse.EntityMetadata
                .Where(
                    entityMetadata =>
                        string.Equals(entitySelected, entityMetadata.LogicalName, StringComparison.OrdinalIgnoreCase))
                .Select(entityMetadata => entityMetadata.ManyToOneRelationships).ToList().FirstOrDefault();

            var oneToManyAttributes = _metadataResponse.EntityMetadata
                .Where(
                    entityMetadata =>
                        string.Equals(entitySelected, entityMetadata.LogicalName, StringComparison.OrdinalIgnoreCase))
                .Select(entityMetadata => entityMetadata.OneToManyRelationships).ToList().FirstOrDefault();

            if (oneToManyAttributes == null)
                return null;

            var output = oneToManyAttributes
                    .Select(attribute => attribute.ReferencingEntity + " - " + attribute.ReferencingAttribute)
                    .ToList();
            output.Sort();
            return output;
        }

        internal OneToManyRelationshipMetadata GetRelationShipByReferencingAttribute(string attributeName)
        {
            if (_entityRelationshipMetadataAttributeList == null)
                return null;

            return _entityRelationshipMetadataAttributeList
                .FirstOrDefault(x => x.ReferencingAttribute == attributeName);
        }
    }
}
