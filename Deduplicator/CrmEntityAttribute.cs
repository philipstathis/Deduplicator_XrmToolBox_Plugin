using System;

namespace Deduplicator
{
    internal class CrmEntityAttribute
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public CrmEntityAttribute(string attributeName, string attributeType)
        {
            Name = attributeName;
            Type = attributeType;
        }

        public bool IsLookup
        {
            get { return Type.Contains("Lookup"); }
        }

        public bool IsPrimaryKey
        {
            get { return string.Equals(Type, "Uniqueidentifier", StringComparison.OrdinalIgnoreCase); }
        }

        public bool IsDate
        {
            get { return string.Equals(Type, "DateTime", StringComparison.OrdinalIgnoreCase); }
        }

    }
}
