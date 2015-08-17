using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace Deduplicator
{
    class FormEntryReader
    {
        private readonly FetchXmlBuilder _fetchXmlBuilder;
        private readonly List<CrmEntityAttribute> _uniqueIdentifierAttributes;

        public FormEntryReader(FetchXmlBuilder fetchXmlBuilder)
        {
            _fetchXmlBuilder = fetchXmlBuilder;
            _uniqueIdentifierAttributes = new List<CrmEntityAttribute>();
        }

        private DataRow ExtractItemFromDataRow(object dataRow)
        {
            var row = dataRow as DataGridViewRow;
            if (row != null)
            {
                var item = row.DataBoundItem as DataRowView;
                if (item != null)
                {
                    return item.Row;
                }
            }
            return null;
        }

        internal string ReadEntitySelected(ComboBox entityDropdown)
        {
            var entitySelected = entityDropdown.SelectedValue as string;
            if (string.IsNullOrWhiteSpace(entitySelected))
                throw new ArgumentNullException(entitySelected);

            if (_fetchXmlBuilder != null)
                _fetchXmlBuilder.SetEntity(entitySelected);

            return entitySelected;
        }

        internal void ReadFilterValuesFromSelectedRow(DataGridView gridView)
        {
            var item = GetSelectedRowItem(gridView);
            if (item == null)
                return;

            var uniqueAttributeValuePairs = _uniqueIdentifierAttributes
                .ToDictionary(uniqueAttribute => uniqueAttribute,
                    uniqueAttribute => item[uniqueAttribute.Name]);

            var entityQueryBuilder = _fetchXmlBuilder as EntityQueryBuilder;
            if (entityQueryBuilder != null)
                entityQueryBuilder.AddIdentifierValuePairs(uniqueAttributeValuePairs);

            var relatedEntityQueryBuilder = _fetchXmlBuilder as RelatedEntityQueryBuilder;
            if (relatedEntityQueryBuilder != null)
                relatedEntityQueryBuilder.AddIdentifierValuePairs(uniqueAttributeValuePairs);
        }

        internal DataRow GetSelectedRowItem(DataGridView gridView)
        {
            var selectedRow = gridView.SelectedRows
                .OfType<DataGridViewRow>()
                .FirstOrDefault();

            var item = ExtractItemFromDataRow(selectedRow);
            return item;
        }

        internal void ReadDisplayAttributes(DataGridView entityAttributeView)
        {
            // parse selected columns from DataGridView
            foreach (var dataRow in entityAttributeView.Rows)
            {
                var item = ExtractItemFromDataRow(dataRow);
                if (item == null)
                    return;

                var userSelectedDisplayAttribute = (bool) item["Display in Step 5?"];
                var isPrimaryKey = string.Equals(item["Attribute Type"] as string, "Uniqueidentifier",
                    StringComparison.OrdinalIgnoreCase);

                var addToQuery = isPrimaryKey || userSelectedDisplayAttribute;
                if (!addToQuery)
                    continue;

                var attributeName = item["Logical Name"] as string;
                var attributeType = item["Attribute Type"] as string;
                _fetchXmlBuilder.AddColumn(new CrmEntityAttribute(attributeName, attributeType));
            }
        }

        internal void ReadUniqueIdentifierAttributes(DataGridView entityAttributeView)
        {
            // parse selected columns from DataGridView
            foreach (var dataRow in entityAttributeView.Rows)
            {
                var item = ExtractItemFromDataRow(dataRow);

                var addToQuery = (bool)item["Set As Unique Identifier?"];
                if (!addToQuery)
                    continue;

                var attributeName = item["Logical Name"] as string;
                var attributeType = item["Attribute Type"] as string;

                var attribute = new CrmEntityAttribute(attributeName, attributeType);
                _uniqueIdentifierAttributes.Add(attribute);
            }

            // if we are building a DuplicateQuery we can simply use the identifiers right away
            if (!(_fetchXmlBuilder is DuplicateQueryBuilder)) return;

            foreach (var attribute in _uniqueIdentifierAttributes)
            {
                _fetchXmlBuilder.AddColumn(attribute);
            }
        }

        internal void ReadRelatedEntitySelected(ComboBox relationShipDropdown)
        {
            var relationshipString = relationShipDropdown.SelectedValue as string;

            if (relationshipString == null) return;
            var entries = relationshipString.Split(new[] {" - "},StringSplitOptions.None);

            ReferencingEntityName = entries[0];
            ReferencingAttributeName = entries[1];

            var entityQueryBuilder = _fetchXmlBuilder as RelatedEntityQueryBuilder;

            if (entityQueryBuilder == null) return;

            entityQueryBuilder.SetRelatedEntity(ReferencingEntityName, ReferencingAttributeName);
        }

        public string ReferencingEntityName { get; private set; }

        public string ReferencingAttributeName { get; private set; }
    }
}
