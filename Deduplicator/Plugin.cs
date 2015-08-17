using System;
using System.ComponentModel.Composition;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace Deduplicator
{
    [Export(typeof(IXrmToolBoxPlugin)),
      ExportMetadata("Name", "Deduplicator"),
      ExportMetadata("Description", "Merge/Delete tool"),
      ExportMetadata("SmallImageBase64", null), // null for "no logo" image or base64 image content 
      ExportMetadata("BigImageBase64", null), // null for "no logo" image or base64 image content 
      ExportMetadata("BackgroundColor", "Lavender"), // Use a HTML color name
      ExportMetadata("PrimaryFontColor", "#000000"), // Or an hexadecimal code
      ExportMetadata("SecondaryFontColor", "DarkGray")]

    internal class Plugin : PluginBase
    {
        public override IXrmToolBoxPluginControl GetControl()
        {
            return new Main();
        }
    }
}