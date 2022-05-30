using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Data;
using GHD.ContractRenewal.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;

namespace GHD.ContractRenewal.Activities
{
    [LocalizedDisplayName(nameof(Resources.GenerateContract_DisplayName))]
    [LocalizedDescription(nameof(Resources.GenerateContract_Description))]
    public class GenerateContract : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.GenerateContract_ExceptionMessage_DisplayName))]
        [LocalizedDescription(nameof(Resources.GenerateContract_ExceptionMessage_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> ExceptionMessage { get; set; }

        [LocalizedDisplayName(nameof(Resources.GenerateContract_WordOutputPath_DisplayName))]
        [LocalizedDescription(nameof(Resources.GenerateContract_WordOutputPath_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> WordOutputPath { get; set; }

        [LocalizedDisplayName(nameof(Resources.GenerateContract_LogDirectory_DisplayName))]
        [LocalizedDescription(nameof(Resources.GenerateContract_LogDirectory_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> LogDirectory { get; set; }

        [LocalizedDisplayName(nameof(Resources.GenerateContract_NonBidTemplatePath_DisplayName))]
        [LocalizedDescription(nameof(Resources.GenerateContract_NonBidTemplatePath_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> NonBidTemplatePath { get; set; }

        [LocalizedDisplayName(nameof(Resources.GenerateContract_BidTemplatePath_DisplayName))]
        [LocalizedDescription(nameof(Resources.GenerateContract_BidTemplatePath_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> BidTemplatePath { get; set; }

        [LocalizedDisplayName(nameof(Resources.GenerateContract_OutputFolderPath_DisplayName))]
        [LocalizedDescription(nameof(Resources.GenerateContract_OutputFolderPath_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> OutputFolderPath { get; set; }

        [LocalizedDisplayName(nameof(Resources.GenerateContract_DataTableInfo_DisplayName))]
        [LocalizedDescription(nameof(Resources.GenerateContract_DataTableInfo_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<DataTable> DataTableInfo { get; set; }

        [LocalizedDisplayName(nameof(Resources.GenerateContract_DataTableAnnual_DisplayName))]
        [LocalizedDescription(nameof(Resources.GenerateContract_DataTableAnnual_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<DataTable> DataTableAnnual { get; set; }

        [LocalizedDisplayName(nameof(Resources.GenerateContract_DataTableOptional_DisplayName))]
        [LocalizedDescription(nameof(Resources.GenerateContract_DataTableOptional_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<DataTable> DataTableOptional { get; set; }

        #endregion


        #region Constructors

        public GenerateContract()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var logDirectory = LogDirectory.Get(context);
            var nonBidTemplatePath = NonBidTemplatePath.Get(context);
            var bidTemplatePath = BidTemplatePath.Get(context);
            var outputFolderPath = OutputFolderPath.Get(context);
            var dataTableInfo = DataTableInfo.Get(context);
            var dataTableAnnual = DataTableAnnual.Get(context);
            var dataTableOptional = DataTableOptional.Get(context);
    
            ///////////////////////////
            // Add execution logic HERE
            ///////////////////////////

            // Outputs
            return (ctx) => {
                ExceptionMessage.Set(ctx, null);
                WordOutputPath.Set(ctx, null);
            };
        }

        #endregion
    }
}

