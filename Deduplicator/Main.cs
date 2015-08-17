using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using XrmToolBox.Extensibility;

namespace Deduplicator
{
    public partial class Main : PluginControlBase
    {
        public Main()
        {
            InitializeComponent();
        }

        private void queryEntities_Click(object sender, EventArgs e)
        {
            ExecuteMethod(QueryEntityMetadata);
        }

        private void queryAttributes_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(entityDropdown.SelectedValue as string))
            {
                MessageBox.Show("Please select entity from step 1");
                return;
            }
            ExecuteMethod(QueryAttributeMetadata);
        }

        private void queryRelationships_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(entityDropdown.SelectedValue as string))
            {
                MessageBox.Show("Please select entity from step 1");
                return;
            }
            ExecuteMethod(QueryRelationshipMetadata);
        }

        private void queryDuplicates_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(entityDropdown.SelectedValue as string))
            {
                MessageBox.Show("Please select entity from step 1");
                return;
            }
            // check that columns are selected
            ExecuteMethod(QueryDuplicates);
        }

        private void querySelectedDuplicate_Click(object sender, EventArgs e)
        {
            ExecuteMethod(QueryRecords);
        }

        private void updateBtn_Click(object sender, EventArgs e)
        {
            ExecuteMethod(UpdateRelatedRecords);
        }

        private void deleteBtn_Click(object sender, EventArgs e)
        {
            ExecuteMethod(DeleteMarkedRecords);
        }

        private void queryRelatedEntitiesBtn_Click(object sender, EventArgs e)
        {
            ExecuteMethod(QueryRelatedRecords);
        }

        private void UpdateRelatedRecords()
        {
            WorkAsync("Updating Related Records...",
                (w, e) => // Work To Do Asynchronously
                {
                    w.ReportProgress(0, "Parsing Selected Surviving Record");

                    // Uncheck Selected Record from the To Be Deleted List
                    var formReader = new FormEntryReader(null);
                    var selectedEntityRecordFromStep5 = formReader.GetSelectedRowItem(entityRecordGrid);
                    selectedEntityRecordFromStep5["Mark for Deletion"] = false;
                    entityRecordGrid.Refresh();

                    // Get RelatedEntityValues
                    var entitySelected = formReader.ReadEntitySelected(entityDropdown);
                    formReader.ReadRelatedEntitySelected(relationShipDropdown);
                    var relatedEntityPkName = formReader.ReferencingEntityName + "id";
                    var relatedEntity = formReader.ReferencingEntityName;
                    var linkColumnOnRelatedEntity = formReader.ReferencingAttributeName;

                    var formUpdater = new GridUpdater(relatedEntityGrid);
                    formUpdater.Where(row => row[relatedEntityPkName] is Guid)
                        .ForEach(row =>
                        {
                            var currentRow = row.Table.Rows.IndexOf(row) + 1;
                            // TO DO: The indexing totals system can be improved to cover actual work rows
                            var totalRows = row.Table.Rows.Count;
                            w.ReportProgress(currentRow / totalRows, "Updating Row ID: " + currentRow);
                            var updateItem = new Entity(relatedEntity) {Id = (Guid) row[relatedEntityPkName]};
                            updateItem[linkColumnOnRelatedEntity] = new EntityReference(
                                entitySelected,
                                (Guid)selectedEntityRecordFromStep5[entitySelected + "id"]);
                            Service.Update(updateItem);
                            row["Updated"] = true;
                        });
                    w.ReportProgress(99, "Finishing");
                    formUpdater.Refresh();

                    e.Result = "Update Complete";
                },
                e => // Finished Async Call.  Cleanup
                    MessageBox.Show(e.Result as string),
                e => // Logic wants to display an update.  This gets called when ReportProgress Gets Called
                    SetWorkingMessage(e.UserState.ToString()));
        }

        private void QueryRelatedRecords()
        {
            WorkAsync("Querying Related Entities for Selected Row from Step 4",
                e => // Work To Do Asynchronously
                {
                    var fetchXmlBuilder = new RelatedEntityQueryBuilder();
                    var formEntryReader = new FormEntryReader(fetchXmlBuilder);
                    formEntryReader.ReadEntitySelected(entityDropdown);
                    formEntryReader.ReadUniqueIdentifierAttributes(entityAttributeView);
                    formEntryReader.ReadFilterValuesFromSelectedRow(duplicatesGrid);
                    formEntryReader.ReadRelatedEntitySelected(relationShipDropdown);
                    fetchXmlBuilder.AddColumn(new CrmEntityAttribute("createdon", "DateTime"));
                    var response = GetRecordsFromCrm(fetchXmlBuilder.GetOutput());

                    var reader = new FetchXmlResponseReader(response);
                    var dataset = reader.ParseRelatedRecords(fetchXmlBuilder);
                    SafeSetDataSource(relatedEntityGrid, dataset);

                    e.Result = "Step 6 Record View Updated";
                }, e =>
                {
                    MessageBox.Show(e.Result as string);
                });
        }

        private void DeleteMarkedRecords()
        {
            WorkAsync("Deleting Marked Records...",
                (w, e) => // Work To Do Asynchronously
                {
                    var formReader = new FormEntryReader(null);
                    var entitySelected = formReader.ReadEntitySelected(entityDropdown);

                    w.ReportProgress(0, "Start Parsing");
                    var formUpdater = new GridUpdater(entityRecordGrid);
                    formUpdater.Where(row => (bool) row["Mark for Deletion"])
                        .ForEach(row =>
                        {
                            var currentRow = row.Table.Rows.IndexOf(row) + 1;
                            // TO DO: The indexing totals system can be improved to cover actual work rows
                            var totalRows = row.Table.Rows.Count;
                            w.ReportProgress(currentRow / totalRows, "Deleting Row ID: " + currentRow);
                            Service.Delete(entitySelected, (Guid)row[entitySelected+"id"]);
                        });
                    w.ReportProgress(99, "Finishing");
                    formUpdater.Refresh();

                    // Populate whatever the results that need to be returned to the Results Property
                    e.Result = "Deletion Complete";
                },
                e => // Finished Async Call.  Cleanup
                    MessageBox.Show(e.Result as string),
                e => // Logic wants to display an update.  This gets called when ReportProgress Gets Called
                    SetWorkingMessage(e.UserState.ToString()));
        }

        private void QueryRecords()
        {
            WorkAsync("Querying Records for Selected Rows from Step 4",
                e => // Work To Do Asynchronously
                {
                    var fetchXmlBuilder = new EntityQueryBuilder();
                    var formEntryReader = new FormEntryReader(fetchXmlBuilder);
                    formEntryReader.ReadEntitySelected(entityDropdown);
                    formEntryReader.ReadUniqueIdentifierAttributes(entityAttributeView);
                    formEntryReader.ReadFilterValuesFromSelectedRow(duplicatesGrid);
                    formEntryReader.ReadDisplayAttributes(entityAttributeView);
                    var response = GetRecordsFromCrm(fetchXmlBuilder.GetOutput());

                    var reader = new FetchXmlResponseReader(response);
                    var dataset = reader.ParseEntityQueryResponse(fetchXmlBuilder);
                    SafeSetDataSource(entityRecordGrid, dataset);

                    e.Result = "Step 5 Record View Updated";
                }, e => MessageBox.Show(e.Result as string));
        }

        private void QueryDuplicates()
        {
            WorkAsync(
                string.Format("Querying Duplicates [Entity:{0} - Attributes Selected:{1}", entityDropdown.SelectedValue,
                    entityAttributeView.RowCount),
                e => // Work To Do Asynchronously
                {
                    var fetchXmlBuilder = new DuplicateQueryBuilder();
                    var formEntryReader = new FormEntryReader(fetchXmlBuilder);
                    formEntryReader.ReadEntitySelected(entityDropdown);
                    formEntryReader.ReadUniqueIdentifierAttributes(entityAttributeView);
                    fetchXmlBuilder.AddCustomFilterXml(customXMLFilterBox.Text);
                    e.Result = GetRecordsFromCrm(fetchXmlBuilder.GetOutput());

                },
                e => // Cleanup when work has completed
                {
                    var reader = new FetchXmlResponseReader(e.Result as EntityCollection);

                    var fetchXmlBuilder = new DuplicateQueryBuilder();
                    var formEntryReader = new FormEntryReader(fetchXmlBuilder);
                    formEntryReader.ReadEntitySelected(entityDropdown);
                    formEntryReader.ReadUniqueIdentifierAttributes(entityAttributeView);
                    fetchXmlBuilder.AddCustomFilterXml(customXMLFilterBox.Text);
                    var dataSet = reader.ParseGroupByResponse(fetchXmlBuilder);
                    SafeSetDataSource(duplicatesGrid, dataSet);
                });
        }

        private void QueryRelationshipMetadata()
        {
            WorkAsync("Querying Relationships for Selected Entity: " + entityDropdown.SelectedValue,
                e => // Work To Do Asynchronously
                {
                    var request = new RetrieveAllEntitiesRequest
                    {
                        EntityFilters = EntityFilters.Relationships,
                        RetrieveAsIfPublished = true
                    };

                   e.Result = Service.Execute(request);
                }, e =>
                {
                    var reader = new EntityMetadataReader(e.Result as OrganizationResponse);
                    relationShipDropdown.DataSource = reader.ParseRelationships(entityDropdown.SelectedValue as string);

                    var formUpdater = new GridUpdater(entityAttributeView)
                    {
                        ColumnForAttributeName = "Logical Name",
                        ColumnForAttributeType = "Attribute Type"
                    };
                    formUpdater.Where(row => row[formUpdater.ColumnForAttributeType] as string == "Lookup")
                        .ForEach(x => formUpdater.OverwriteAttributeTypeWithRelationshipInfo(x, reader));
                    formUpdater.Refresh();
                });
        }

        private void QueryAttributeMetadata()
        {
            WorkAsync("Querying Attributes for Selected Entity: " + entityDropdown.SelectedValue,
                e => // Work To Do Asynchronously
                {
                    var request = new RetrieveAllEntitiesRequest
                    {
                        EntityFilters = EntityFilters.Attributes,
                        RetrieveAsIfPublished = true
                    };

                    var response = Service.Execute(request);

                    var reader = new EntityMetadataReader(response);
                    e.Result = reader.ParseAttributes(entityDropdown.SelectedValue as string);
                },
                e => // Cleanup when work has completed
                    SafeSetDataSource(entityAttributeView, e.Result));
        }

        private void QueryEntityMetadata()
        {
            WorkAsync("Querying Entity List",
                e => // Work To Do Asynchronously
                {
                    var request = new RetrieveAllEntitiesRequest
                    {
                        EntityFilters = EntityFilters.Entity,
                        RetrieveAsIfPublished = true
                    };

                    e.Result = Service.Execute(request);
                }, e =>
                {
                    var reader = new EntityMetadataReader(e.Result as OrganizationResponse);
                    entityDropdown.DataSource = reader.ParseEntityList();
                });
        }

        private void SafeSetDataSource(DataGridView view, object dataSet)
        {
            view.Visible = false;
            view.DataSource = dataSet;
            foreach (DataGridViewColumn c in view.Columns)
            {
                c.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            }
            view.Visible = true;
            //view.Refresh();
        }

        /// <summary>
        /// Retrieves multiple records from CRM. Can retrieve more than 5,000 records through paging.
        /// </summary>
        /// <returns>CRM Records</returns>
        public EntityCollection GetRecordsFromCrm(FetchExpression fetchXml)
        {
            // Define the fetch attributes.
            // Set the number of records per page to retrieve.
            int fetchCount = 1000;
            // Initialize the page number.
            int pageNumber = 1;
            // Initialize the number of records.
            int recordCount = 0;
            // Specify the current paging cookie. For retrieving the first page, 
            // pagingCookie should be null.
            string pagingCookie = null;
            EntityCollection results = new EntityCollection();
            EntityCollection returnRecords = new EntityCollection();
            do
            {
                // Build fetchXml string with the placeholders.
                string xml = CreateXml(fetchXml.Query, pagingCookie, pageNumber, fetchCount);

                results = Service.RetrieveMultiple(new FetchExpression(xml));
                returnRecords.Entities.AddRange(results.Entities);

                // Increment the page number to retrieve the next page.
                pageNumber++;

                // Set the paging cookie to the paging cookie returned from current results.                            
                pagingCookie = returnRecords.PagingCookie;
            } while (results.MoreRecords);
            return returnRecords;
        }

        public string ExtractNodeValue(XmlNode parentNode, string name)
        {
            XmlNode childNode = parentNode.SelectSingleNode(name);

            if (null == childNode)
            {
                return null;
            }
            return childNode.InnerText;
        }

        public string ExtractAttribute(XmlDocument doc, string name)
        {
            if (doc.DocumentElement == null)
                throw new ArgumentNullException("DocumentElement");

            XmlAttributeCollection attrs = doc.DocumentElement.Attributes;
            XmlAttribute attr = (XmlAttribute) attrs.GetNamedItem(name);
            if (null == attr)
            {
                return null;
            }
            return attr.Value;
        }

        public string CreateXml(string xml, string cookie, int page, int count)
        {
            StringReader stringReader = new StringReader(xml);
            XmlTextReader reader = new XmlTextReader(stringReader);

            // Load document
            XmlDocument doc = new XmlDocument();
            doc.Load(reader);

            return CreateXml(doc, cookie, page, count);
        }

        public string CreateXml(XmlDocument doc, string cookie, int page, int count)
        {
            if (doc.DocumentElement == null)
                throw new ArgumentNullException("DocumentElement");

            XmlAttributeCollection attrs = doc.DocumentElement.Attributes;

            if (cookie != null)
            {
                XmlAttribute pagingAttr = doc.CreateAttribute("paging-cookie");
                pagingAttr.Value = cookie;
                attrs.Append(pagingAttr);
            }

            XmlAttribute pageAttr = doc.CreateAttribute("page");
            pageAttr.Value = System.Convert.ToString(page);
            attrs.Append(pageAttr);

            XmlAttribute countAttr = doc.CreateAttribute("count");
            countAttr.Value = System.Convert.ToString(count);
            attrs.Append(countAttr);

            StringBuilder sb = new StringBuilder(1024);
            StringWriter stringWriter = new StringWriter(sb);

            XmlTextWriter writer = new XmlTextWriter(stringWriter);
            doc.WriteTo(writer);
            writer.Close();

            return sb.ToString();
        }

        private void Output_Debug(object sender, EventArgs e)
        {
            var fetchXmlBuilder = new DuplicateQueryBuilder();
            var formEntryReader = new FormEntryReader(fetchXmlBuilder);
            formEntryReader.ReadEntitySelected(entityDropdown);
            formEntryReader.ReadUniqueIdentifierAttributes(entityAttributeView);
            fetchXmlBuilder.AddCustomFilterXml(customXMLFilterBox.Text);
            var query = fetchXmlBuilder.GetOutput().Query;
            MessageBox.Show(string.Format("Hit CTRL - C to copy the query below: {0}", query));
        }

        private void ClosePlugin_Click(object sender, EventArgs e)
        {
            CloseTool();
        }
    }
}
