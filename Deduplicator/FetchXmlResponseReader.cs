using System;
using System.Data;
using System.Globalization;
using Microsoft.Xrm.Sdk;

namespace Deduplicator
{
    internal class FetchXmlResponseReader : CrmReader
    {
        private readonly EntityCollection _response;

        public FetchXmlResponseReader(EntityCollection response)
        {
            _response = response;
        }

        internal object ParseGroupByResponse(DuplicateQueryBuilder schema)
        {
            var sampleDataSet = new DataSet {Locale = CultureInfo.InvariantCulture};
            var output = sampleDataSet.Tables.Add("Duplicates");
            output.Columns.Add(schema.DuplicateColumnName, typeof(int));

            SetTableColumnsFromSchema(schema, output);

            foreach (var response in _response.Entities)
            {
                var countValue = GetAliasedValue(response, schema.DuplicateColumnName);

                if ((int) countValue == 1)
                    continue;

                var row = output.NewRow();
                row[schema.DuplicateColumnName] = countValue;

                foreach (var attribute in schema.Attributes)
                {
                    if (attribute.IsDate)
                    {
                        var groupedDay = GetAliasedValue(response, "GroupByDay");
                        var groupedMonth = GetAliasedValue(response, "GroupByMonth");
                        var groupedYear = GetAliasedValue(response, "GroupByYear");
                        row[attribute.Name] = groupedMonth + "/" + groupedDay + "/" + groupedYear;
                    }
                    else if (attribute.IsLookup)
                    {
                        row[attribute.Name] = GetAliasedValue(response, "GroupBy" + attribute.Name);
                        row[attribute.Name + "name"] = GetAliasedName(response, "GroupBy" + attribute.Name);
                    }
                    else
                    {
                        row[attribute.Name] = GetAliasedValue(response, "GroupBy" + attribute.Name);
                    }
                }
                output.Rows.Add(row);
            }

            output.AcceptChanges();
            return output;
        }

        internal object ParseEntityQueryResponse(EntityQueryBuilder schema)
        {
            var sampleDataSet = new DataSet { Locale = CultureInfo.InvariantCulture };
            DataTable output = sampleDataSet.Tables.Add("Entity");

            SetTableColumnsFromSchema(schema, output);

            output.Columns.Add("Mark for Deletion", typeof (bool));

            foreach (var response in _response.Entities)
            {
                var row = output.NewRow();

                ParseResponseUsingSchema(schema, row, response);
                row["Mark for Deletion"] = true;
                output.Rows.Add(row);
            }

            output.AcceptChanges();
            return output;
        }

        internal object ParseRelatedRecords(RelatedEntityQueryBuilder schema)
        {
            var sampleDataSet = new DataSet { Locale = CultureInfo.InvariantCulture };
            DataTable output = sampleDataSet.Tables.Add("Related Entities");

            SetTableColumnsFromSchema(schema, output);

            output.Columns.Add("Updated", typeof(bool));

            foreach (var response in _response.Entities)
            {
                var row = output.NewRow();
                ParseResponseUsingSchema(schema, row, response);
                output.Rows.Add(row);
            }

            output.AcceptChanges();
            return output;
        }

        private static void SetTableColumnsFromSchema(FetchXmlBuilder schema, DataTable output)
        {
            foreach (var attribute in schema.Attributes)
            {
                if (attribute.IsLookup)
                {
                    output.Columns.Add(attribute.Name, typeof (Guid));
                    output.Columns.Add(attribute.Name + "name", typeof (string));
                }
                else if (attribute.IsPrimaryKey)
                    output.Columns.Add(attribute.Name, typeof (Guid));
                else if (attribute.IsDate)
                    output.Columns.Add(attribute.Name, typeof (DateTime));
                else
                    output.Columns.Add(attribute.Name, typeof (string));
            }
        }

        private void ParseResponseUsingSchema(FetchXmlBuilder schema, DataRow row, Entity response)
        {
            foreach (var attribute in schema.Attributes)
            {
                if (attribute.IsLookup)
                {
                    row[attribute.Name] = GetEntityReferenceValue(response, attribute.Name);
                    row[attribute.Name + "name"] = GetEntityReferenceName(response, attribute.Name);
                }
                else
                {
                    row[attribute.Name] = GetValue(response, attribute.Name);
                }
            }
        }
    }
}
