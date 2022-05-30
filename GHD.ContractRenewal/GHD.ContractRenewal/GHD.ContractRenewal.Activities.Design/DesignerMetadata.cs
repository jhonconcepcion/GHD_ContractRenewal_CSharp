using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using GHD.ContractRenewal.Activities.Design.Designers;
using GHD.ContractRenewal.Activities.Design.Properties;

namespace GHD.ContractRenewal.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(GenerateContract), categoryAttribute);
            builder.AddCustomAttributes(typeof(GenerateContract), new DesignerAttribute(typeof(GenerateContractDesigner)));
            builder.AddCustomAttributes(typeof(GenerateContract), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
