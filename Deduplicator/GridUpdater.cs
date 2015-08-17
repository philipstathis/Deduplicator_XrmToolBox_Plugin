using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Xrm.Sdk.Metadata;

namespace Deduplicator
{
    class GridUpdater
    {
        private readonly DataGridView _entityAttributeView;

        public GridUpdater(DataGridView entityAttributeView)
        {
            _entityAttributeView = entityAttributeView;
        }

        internal List<DataRow> Where(Func<DataRow, bool> func)
        {
            return _entityAttributeView.Rows
                .OfType<DataGridViewRow>()
                .Select(viewItem => viewItem.DataBoundItem)
                .OfType<DataRowView>()
                .Select(item => item.Row)
                .Where(func).ToList();
        }

        public string ColumnForAttributeName { get; set; }

        public string ColumnForAttributeType { get; set; }

        internal void OverwriteAttributeTypeWithRelationshipInfo(DataRow row, EntityMetadataReader reader)
        {
            var attributeName = row[ColumnForAttributeName] as string;

            var metadata = reader.GetRelationShipByReferencingAttribute(attributeName);

            if (metadata == null)
                return;

            row[ColumnForAttributeType] = string.Format("Lookup[RefEntity:{0}][RefColumnName:{1}]",
                metadata.ReferencedEntity, metadata.ReferencedAttribute);
        }

        internal void Refresh()
        {
            _entityAttributeView.Refresh();
        }
    }
}
