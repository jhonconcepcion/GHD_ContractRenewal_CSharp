using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Data;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using GHD.ContractRenewal.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

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

            //Output
            string strWordOutput = "";
            string strExceptionMessage = "";
            
            


            try
            {
                UpdateWordDoc(dataTableInfo);
            }
            catch(Exception e)
            {
                strExceptionMessage = e.Message;
            }
           
            void UpdateWordDoc(DataTable dtInfo)
           {
                Word.Application app = null;
                Word.Documents docs = null;
                Word.Document doc = null;
                Word.Table tblOrderTerms = null;
                Word.Table tblSOW = null;
                Word.Table tblOptional = null;
                bool IsWordAppLaunched = false;
                bool isBid = getDocType(dtInfo);
                string strFileName = getFileName(dtInfo);
                object missing = System.Reflection.Missing.Value;
                Dictionary<string, string> dicInfo = buildCustomerInfo(dtInfo);


                try
                {
                    Log("UpdateWordDoc" + " Method name:" + MethodBase.GetCurrentMethod().Name + " " + "UpdateWordDoc function start");
                    app = new Word.Application();
                    app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                    IsWordAppLaunched = true;
                    docs = app.Documents;
                    Log("UpdateWordDoc" + " Method name:" + MethodBase.GetCurrentMethod().Name + " " + "Word application object initialized");
                    string newfolderName = DateTime.Now.ToString("MMddyyyy HHmmss");
                    string newFilePath = CreateNewWordDoc(newfolderName, isBid, strFileName);
                    doc = docs.Open(newFilePath, ReadOnly: false);
                    Log("Opened word document");

                    tblSOW = doc.Tables[2];
                    tblOptional = doc.Tables[3];

                    foreach (KeyValuePair<string, string> item in dicInfo)
                    {
                        /* This part loops through dicInfo and uses the key as the item to be searched in the word template.
                         * 
                         * app.Application.Selection.Find.Text = "[" + item.Key + "]"; this part searches to the word or item to be replaced.
                         * The if and else part are the code that replaces the placeholder with the actual value.
                         * The if part is for values with less than 256 in character length and the else part is for items equals or greater than 256;
                         * 
                         */
                        object replaceAll = Word.WdReplace.wdReplaceAll;
                        app.Application.Selection.Find.ClearFormatting();
                        app.Application.Selection.Find.Text = "[" + item.Key + "]";
                        app.Application.Selection.Find.Replacement.ClearFormatting();
                        string replaceTxt = item.Value.Trim();
                        if (replaceTxt.ToString().Length < 256)
                        {
                            app.Application.Selection.Find.Replacement.Text = replaceTxt;
                            bool IsReplaced = app.Application.Selection.Find.Execute(
                            ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref replaceAll, ref missing, ref missing, ref missing, ref missing);
                        }
                        else
                        {
                            object findme = "[" + item.Key + "]";
                            while (app.Application.Selection.Find.Execute(
                                  ref findme, ref missing, ref missing, ref missing, ref missing,
                                  ref missing, ref missing, ref missing, ref missing, ref missing,
                                  ref missing, ref missing, ref missing, ref missing, ref missing))
                            {

                                app.Application.Selection.Text = replaceTxt;
                                app.Application.Selection.Collapse();
                            }
                        }
                    }

                    
                    /* Update SOW Table
                     * cell(i+2) the +2 is for the buffer tbl row starts at 1, the 1st row is the header so it should starts at index 2
                     
                    for (int i = 0; i< dtAnnual.Rows.Count; i++)
                    {
                        for(int j = 0; j < dtAnnual.Columns.Count; j++)
                        {
                            tblSOW.Cell(i + 2, j).Range.Text = dtAnnual.Rows[i][j].ToString();
                        }
                    }

                    //Update Optional table
                    for (int i = 0; i < dtOptional.Rows.Count; i++)
                    {
                        for (int j = 0; j < dtAnnual.Columns.Count; j++)
                        {
                            tblOptional.Cell(i + 2, j).Range.Text = dtOptional.Rows[i][j].ToString();
                        }
                    }
                    */
                    doc.Save();
                    doc.Close(ref missing, ref missing, ref missing);
                    doc = null;
                    Log("Document saved and closed successfully");

                }
                catch(Exception ex)
                {
                    strExceptionMessage = "Exception occured on generating the word document, this might be an issue with the supporting file, please check if the stockpile names are correct and please make sure that there is no merged cells." + "Error Code: " + ex.Message;
                    Log("Exception Occured: " + ex + " occured in " + MethodBase.GetCurrentMethod().Name + "Error Code: " + ex.Message);



                    if (IsWordAppLaunched)
                    {
                        app.Quit(SaveChanges: true, ref missing, ref missing);
                        app = null;
                    }
                }
                finally
                {
                    if (tblOrderTerms != null) { Marshal.ReleaseComObject(tblOrderTerms); }
                    if (tblSOW != null) { Marshal.ReleaseComObject(tblSOW); }
                    if (tblOptional != null) { Marshal.ReleaseComObject(tblOptional); }
                    if (doc != null) { Marshal.ReleaseComObject(doc); }
                    if (docs != null) { Marshal.ReleaseComObject(docs); }
                    if (app != null) { Marshal.ReleaseComObject(app); }
                }
            }

            Dictionary<string, string> buildCustomerInfo(DataTable dtInfo)
            {
                /*
                 * Building dictionary of items to replace in the word doc template. 
                 * This doesn't include replacing items in Optional and SOW tables.
                 * Key should be the placeholder in the template. Note that the key is case sensitive.
                 */
                var customInfo = new Dictionary<string, string>();
                decimal AnnualFee1;
                decimal AnnualFee2;
                decimal AnnualFee3;
                decimal AnnualFee4;
                decimal AnnualFee5;
                decimal AnnualFee6;
                decimal AnnualFee7;
                decimal AnnualFee8;
                decimal AnnualFee9;
                decimal AnnualFee10;
                decimal AnnualFee11;
                decimal AnnualFee12;
                decimal AnnualFee13;
                decimal AnnualFee14;
                decimal AnnualFee15;
                decimal AnnualFee16;
                decimal AnnualFee17;
                decimal AnnualFee18;
                decimal AnnualFee19;
                decimal AnnualFee20;
                decimal ImplementationFee1;
                decimal ImplementationFee2;
                decimal ImplementationFee3;
                decimal ImplementationFee4;
                decimal ImplementationFee5;
                decimal ImplementationFee6;
                decimal ImplementationFee7;
                decimal ImplementationFee8;
                decimal ImplementationFee9;
                decimal ImplementationFee10;
                decimal ImplementationFee11;
                decimal ImplementationFee12;
                decimal ImplementationFee13;
                decimal ImplementationFee14;
                decimal ImplementationFee15;
                decimal ImplementationFee16;
                decimal ImplementationFee17;
                decimal ImplementationFee18;
                decimal ImplementationFee19;
                decimal ImplementationFee20;
                decimal OptionalAnnualFee1;
                decimal OptionalAnnualFee2;
                decimal OptionalImplementationFee1;
                decimal OptionalImplementationFee2;
                decimal TotalImplementationFee;
                decimal TotalAnnualFee;

                foreach (DataRow drow in dtInfo.Rows)
                {
                    customInfo["Customer Name"] = drow["Customer Name"].ToString();
                    customInfo["Primary Contact"] = drow["Primary Contact"].ToString();
                    customInfo["Title"] = drow["Title"].ToString();
                    customInfo["Email"] = drow["Email"].ToString();
                    customInfo["Telephone"] = drow["Telephone"].ToString();
                    customInfo["Address Street"] = drow["Address Street"].ToString();
                    customInfo["Address City"] = drow["Address City, province"].ToString();
                    customInfo["Address postal code"] = drow["Address postal code"].ToString();
                    customInfo["Send Invoice"] = drow["Send Invoices to"].ToString();
                    customInfo["Product Solution"] = drow["Product Solution"].ToString();
                    customInfo["Project Number"] = drow["Project Number"].ToString();
                    //customInfo["Delivery Timing"] = drow["Estimated Delivery Timing"].ToString();
                    customInfo["License Term"] = drow["License Term"].ToString();
                    customInfo["Payment Terms"] = drow["Payment Terms"].ToString();
                    //customInfo["Expiry Date"] = drow["Quote Expiry Date"].ToString();
                    customInfo["Executive"] = drow["Account Executive"].ToString();
                    customInfo["Executive Email"] = drow["Account Executive Email"].ToString();
                    //customInfo["Executive Phone"] = drow["Account Executive Phone"].ToString();
                    //customInfo["Purpose"] = drow["Purpose"].ToString();

                    var isTotalImplementationFeeDecimal = Decimal.TryParse(drow["Total Implmentation Fee"].ToString(), out TotalImplementationFee);
                    var isTotalAnnualFeeDecimal = Decimal.TryParse(drow["Total Annual Fee"].ToString(), out TotalAnnualFee);
                    if (isTotalImplementationFeeDecimal == true)
                    {
                        customInfo["Total Implementation"] = "$" + String.Format("{0:#,0.00}", drow["Total Implmentation Fee"]);
                    }
                    else
                    {
                        customInfo["Total Implementation"] = drow["Total Implmentation Fee"].ToString();
                    }
                    if (isTotalAnnualFeeDecimal == true)
                    {
                        customInfo["Total Annual"] = "$" + String.Format("{0:#,0.00}", drow["Total Annual Fee"]);
                    }
                    else
                    {
                        customInfo["Total Annual"] = drow["Total Annual Fee"].ToString();
                    }

                    //customInfo["Acceptance Criteria"] = drow["Acceptance Criteria"].ToString();
                    //customInfo["Delivery Schedule"] = drow["Delivery Schedule"].ToString();
                    //customInfo["Exclusions"] = drow["Exclusions and Assumptions"].ToString();
                    //customInfo["Payment Schedule"] = drow["Payment Schedule On-Time"].ToString();
                    //customInfo["Year Fee"] = drow["Payment Schedule Year's fees"].ToString();
                    //customInfo["Sign Name"] = drow["Customer Sign Print Name"].ToString();
                    //customInfo["Sign Title"] = drow["Customer Sign Print Title"].ToString();
                    //customInfo["Sign Date"] = drow["Customer Sign Date"].ToString();
                    //customInfo["GHD Sign Date"] = drow["GHD Digital Sign Date"].ToString();
                    customInfo["Service1"] = drow["Services/Features 1"].ToString();
                    customInfo["Service2"] = drow["Services/Features 2"].ToString();
                    customInfo["Service3"] = drow["Services/Features 3"].ToString();
                    customInfo["Service4"] = drow["Services/Features 4"].ToString();
                    customInfo["Service5"] = drow["Services/Features 5"].ToString();
                    customInfo["Service6"] = drow["Services/Features 6"].ToString();
                    customInfo["Service7"] = drow["Services/Features 7"].ToString();
                    customInfo["Service8"] = drow["Services/Features 8"].ToString();
                    customInfo["Service9"] = drow["Services/Features 9"].ToString();
                    customInfo["Service10"] = drow["Services/Features 10"].ToString();
                    customInfo["Service11"] = drow["Services/Features 11"].ToString();
                    customInfo["Service12"] = drow["Services/Features 12"].ToString();
                    customInfo["Service13"] = drow["Services/Features 13"].ToString();

                    customInfo["Service14"] = drow["Services/Features 14"].ToString();
                    customInfo["Service15"] = drow["Services/Features 15"].ToString();
                    customInfo["Service16"] = drow["Services/Features 16"].ToString();
                    customInfo["Service17"] = drow["Services/Features 17"].ToString();
                    customInfo["Service18"] = drow["Services/Features 18"].ToString();
                    customInfo["Service19"] = drow["Services/Features 19"].ToString();
                    customInfo["Service20"] = drow["Services/Features 20"].ToString();
                    customInfo["Description1"] = drow["Description 1"].ToString();
                    customInfo["Description2"] = drow["Description 2"].ToString();
                    customInfo["Description3"] = drow["Description 3"].ToString();
                    customInfo["Description4"] = drow["Description 4"].ToString();
                    customInfo["Description5"] = drow["Description 5"].ToString();
                    customInfo["Description6"] = drow["Description 6"].ToString();
                    customInfo["Description7"] = drow["Description 7"].ToString();
                    customInfo["Description8"] = drow["Description 8"].ToString();
                    customInfo["Description9"] = drow["Description 9"].ToString();
                    customInfo["Description10"] = drow["Description 10"].ToString();
                    customInfo["Description11"] = drow["Description 11"].ToString();
                    customInfo["Description12"] = drow["Description 12"].ToString();
                    customInfo["Description13"] = drow["Description 13"].ToString();
                    customInfo["Description14"] = drow["Description 14"].ToString();
                    customInfo["Description15"] = drow["Description 15"].ToString();
                    customInfo["Description16"] = drow["Description 16"].ToString();
                    customInfo["Description17"] = drow["Description 17"].ToString();
                    customInfo["Description18"] = drow["Description 18"].ToString();
                    customInfo["Description19"] = drow["Description 19"].ToString();
                    customInfo["Description20"] = drow["Description 20"].ToString();

                    var isImplementationFee1Decimal = Decimal.TryParse(drow["Implementation Fee 1"].ToString(), out ImplementationFee1);
                    var isImplementationFee2Decimal = Decimal.TryParse(drow["Implementation Fee 2"].ToString(), out ImplementationFee2);
                    var isImplementationFee3Decimal = Decimal.TryParse(drow["Implementation Fee 3"].ToString(), out ImplementationFee3);
                    var isImplementationFee4Decimal = Decimal.TryParse(drow["Implementation Fee 4"].ToString(), out ImplementationFee4);
                    var isImplementationFee5Decimal = Decimal.TryParse(drow["Implementation Fee 5"].ToString(), out ImplementationFee5);
                    var isImplementationFee6Decimal = Decimal.TryParse(drow["Implementation Fee 6"].ToString(), out ImplementationFee6);
                    var isImplementationFee7Decimal = Decimal.TryParse(drow["Implementation Fee 7"].ToString(), out ImplementationFee7);
                    var isImplementationFee8Decimal = Decimal.TryParse(drow["Implementation Fee 8"].ToString(), out ImplementationFee8);
                    var isImplementationFee9Decimal = Decimal.TryParse(drow["Implementation Fee 9"].ToString(), out ImplementationFee9);
                    var isImplementationFee10Decimal = Decimal.TryParse(drow["Implementation Fee 10"].ToString(), out ImplementationFee10);
                    var isImplementationFee11Decimal = Decimal.TryParse(drow["Implementation Fee 11"].ToString(), out ImplementationFee11);
                    var isImplementationFee12Decimal = Decimal.TryParse(drow["Implementation Fee 12"].ToString(), out ImplementationFee12);
                    var isImplementationFee13Decimal = Decimal.TryParse(drow["Implementation Fee 13"].ToString(), out ImplementationFee13);
                    var isImplementationFee14Decimal = Decimal.TryParse(drow["Implementation Fee 14"].ToString(), out ImplementationFee14);
                    var isImplementationFee15Decimal = Decimal.TryParse(drow["Implementation Fee 15"].ToString(), out ImplementationFee15);
                    var isImplementationFee16Decimal = Decimal.TryParse(drow["Implementation Fee 16"].ToString(), out ImplementationFee16);
                    var isImplementationFee17Decimal = Decimal.TryParse(drow["Implementation Fee 17"].ToString(), out ImplementationFee17);
                    var isImplementationFee18Decimal = Decimal.TryParse(drow["Implementation Fee 18"].ToString(), out ImplementationFee18);
                    var isImplementationFee19Decimal = Decimal.TryParse(drow["Implementation Fee 19"].ToString(), out ImplementationFee19);
                    var isImplementationFee20Decimal = Decimal.TryParse(drow["Implementation Fee 20"].ToString(), out ImplementationFee20);
                    if (isImplementationFee1Decimal == true)
                    {
                        customInfo["ImplementationFee1"] = "$" + String.Format("{0:#,0.00}", drow["Implementation Fee 1"]);
                    }
                    else
                    {
                        customInfo["ImplementationFee1"] = drow["Implementation Fee 1"].ToString();
                    }
                    if (isImplementationFee2Decimal == true)
                    {
                        customInfo["ImplementationFee2"] = "$" + String.Format("{0:#,0.00}", drow["Implementation Fee 2"]);
                    }
                    else
                    {
                        customInfo["ImplementationFee2"] = drow["Implementation Fee 2"].ToString();
                    }
                    if (isImplementationFee3Decimal == true)
                    {
                        customInfo["ImplementationFee3"] = "$" + String.Format("{0:#,0.00}", drow["Implementation Fee 3"]);
                    }
                    else
                    {
                        customInfo["ImplementationFee3"] = drow["Implementation Fee 3"].ToString();
                    }
                    if (isImplementationFee4Decimal == true)
                    {
                        customInfo["ImplementationFee4"] = "$" + String.Format("{0:#,0.00}", drow["Implementation Fee 4"]);
                    }
                    else
                    {
                        customInfo["ImplementationFee4"] = drow["Implementation Fee 4"].ToString();
                    }
                    if (isImplementationFee5Decimal == true)
                    {
                        customInfo["ImplementationFee5"] = "$" + String.Format("{0:#,0.00}", drow["Implementation Fee 5"]);
                    }
                    else
                    {
                        customInfo["ImplementationFee5"] = drow["Implementation Fee 5"].ToString();
                    }
                    if (isImplementationFee6Decimal == true)
                    {
                        customInfo["ImplementationFee6"] = "$" + String.Format("{0:#,0.00}", drow["Implementation Fee 6"]);
                    }
                    else
                    {
                        customInfo["ImplementationFee6"] = drow["Implementation Fee 6"].ToString();
                    }
                    if (isImplementationFee7Decimal == true)
                    {
                        customInfo["ImplementationFee7"] = "$" + String.Format("{0:#,0.00}", drow["Implementation Fee 7"]);
                    }
                    else
                    {
                        customInfo["ImplementationFee7"] = drow["Implementation Fee 7"].ToString();
                    }
                    if (isImplementationFee8Decimal == true)
                    {
                        customInfo["ImplementationFee8"] = "$" + String.Format("{0:#,0.00}", drow["Implementation Fee 8"]);
                    }
                    else
                    {
                        customInfo["ImplementationFee8"] = drow["Implementation Fee 8"].ToString();
                    }
                    if (isImplementationFee9Decimal == true)
                    {
                        customInfo["ImplementationFee9"] = "$" + String.Format("{0:#,0.00}", drow["Implementation Fee 9"]);
                    }
                    else
                    {
                        customInfo["ImplementationFee9"] = drow["Implementation Fee 9"].ToString();
                    }
                    if (isImplementationFee10Decimal == true)
                    {
                        customInfo["ImplementationFee10"] = "$" + String.Format("{0:#,0.00}", drow["Implementation Fee 10"]);
                    }
                    else
                    {
                        customInfo["ImplementationFee10"] = drow["Implementation Fee 10"].ToString();
                    }
                    if (isImplementationFee11Decimal == true)
                    {
                        customInfo["ImplementationFee11"] = "$" + String.Format("{0:#,0.00}", drow["Implementation Fee 11"]);
                    }
                    else
                    {
                        customInfo["ImplementationFee11"] = drow["Implementation Fee 11"].ToString();
                    }
                    if (isImplementationFee12Decimal == true)
                    {
                        customInfo["ImplementationFee12"] = "$" + String.Format("{0:#,0.00}", drow["Implementation Fee 12"]);
                    }
                    else
                    {
                        customInfo["ImplementationFee12"] = drow["Implementation Fee 12"].ToString();
                    }
                    if (isImplementationFee13Decimal == true)
                    {
                        customInfo["ImplementationFee13"] = "$" + String.Format("{0:#,0.00}", drow["Implementation Fee 13"]);
                    }
                    else
                    {
                        customInfo["ImplementationFee13"] = drow["Implementation Fee 13"].ToString();
                    }
                    if (isImplementationFee14Decimal == true)
                    {
                        customInfo["ImplementationFee14"] = "$" + String.Format("{0:#,0.00}", drow["Implementation Fee 14"]);
                    }
                    else
                    {
                        customInfo["ImplementationFee14"] = drow["Implementation Fee 14"].ToString();
                    }
                    if (isImplementationFee15Decimal == true)
                    {
                        customInfo["ImplementationFee15"] = "$" + String.Format("{0:#,0.00}", drow["Implementation Fee 15"]);
                    }
                    else
                    {
                        customInfo["ImplementationFee15"] = drow["Implementation Fee 15"].ToString();
                    }
                    if (isImplementationFee16Decimal == true)
                    {
                        customInfo["ImplementationFee16"] = "$" + String.Format("{0:#,0.00}", drow["Implementation Fee 16"]);
                    }
                    else
                    {
                        customInfo["ImplementationFee16"] = drow["Implementation Fee 16"].ToString();
                    }
                    if (isImplementationFee17Decimal == true)
                    {
                        customInfo["ImplementationFee17"] = "$" + String.Format("{0:#,0.00}", drow["Implementation Fee 17"]);
                    }
                    else
                    {
                        customInfo["ImplementationFee17"] = drow["Implementation Fee 17"].ToString();
                    }
                    if (isImplementationFee18Decimal == true)
                    {
                        customInfo["ImplementationFee18"] = "$" + String.Format("{0:#,0.00}", drow["Implementation Fee 18"]);
                    }
                    else
                    {
                        customInfo["ImplementationFee18"] = drow["Implementation Fee 18"].ToString();
                    }
                    if (isImplementationFee19Decimal == true)
                    {
                        customInfo["ImplementationFee19"] = "$" + String.Format("{0:#,0.00}", drow["Implementation Fee 19"]);
                    }
                    else
                    {
                        customInfo["ImplementationFee19"] = drow["Implementation Fee 19"].ToString();
                    }
                    if (isImplementationFee20Decimal == true)
                    {
                        customInfo["ImplementationFee20"] = "$" + String.Format("{0:#,0.00}", drow["Implementation Fee 20"]);
                    }
                    else
                    {
                        customInfo["ImplementationFee20"] = drow["Implementation Fee 20"].ToString();
                    }
                    var isAnnualFee1Decimal = Decimal.TryParse(drow["Annual Fee 1"].ToString(), out AnnualFee1);
                    var isAnnualFee2Decimal = Decimal.TryParse(drow["Annual Fee 2"].ToString(), out AnnualFee2);
                    var isAnnualFee3Decimal = Decimal.TryParse(drow["Annual Fee 3"].ToString(), out AnnualFee3);
                    var isAnnualFee4Decimal = Decimal.TryParse(drow["Annual Fee 4"].ToString(), out AnnualFee4);
                    var isAnnualFee5Decimal = Decimal.TryParse(drow["Annual Fee 5"].ToString(), out AnnualFee5);
                    var isAnnualFee6Decimal = Decimal.TryParse(drow["Annual Fee 6"].ToString(), out AnnualFee6);
                    var isAnnualFee7Decimal = Decimal.TryParse(drow["Annual Fee 7"].ToString(), out AnnualFee7);
                    var isAnnualFee8Decimal = Decimal.TryParse(drow["Annual Fee 8"].ToString(), out AnnualFee8);
                    var isAnnualFee9Decimal = Decimal.TryParse(drow["Annual Fee 9"].ToString(), out AnnualFee9);
                    var isAnnualFee10Decimal = Decimal.TryParse(drow["Annual Fee 10"].ToString(), out AnnualFee10);
                    var isAnnualFee11Decimal = Decimal.TryParse(drow["Annual Fee 11"].ToString(), out AnnualFee11);
                    var isAnnualFee12Decimal = Decimal.TryParse(drow["Annual Fee 12"].ToString(), out AnnualFee12);
                    var isAnnualFee13Decimal = Decimal.TryParse(drow["Annual Fee 13"].ToString(), out AnnualFee13);
                    var isAnnualFee14Decimal = Decimal.TryParse(drow["Annual Fee 14"].ToString(), out AnnualFee14);
                    var isAnnualFee15Decimal = Decimal.TryParse(drow["Annual Fee 15"].ToString(), out AnnualFee15);
                    var isAnnualFee16Decimal = Decimal.TryParse(drow["Annual Fee 16"].ToString(), out AnnualFee16);
                    var isAnnualFee17Decimal = Decimal.TryParse(drow["Annual Fee 17"].ToString(), out AnnualFee17);
                    var isAnnualFee18Decimal = Decimal.TryParse(drow["Annual Fee 18"].ToString(), out AnnualFee18);
                    var isAnnualFee19Decimal = Decimal.TryParse(drow["Annual Fee 19"].ToString(), out AnnualFee19);
                    var isAnnualFee20Decimal = Decimal.TryParse(drow["Annual Fee 20"].ToString(), out AnnualFee20);

                    if (isAnnualFee1Decimal == true)
                    {
                        customInfo["AnnualFee1"] = "$" + String.Format("{0:#,0.00}", drow["Annual Fee 1"]);
                    }
                    else
                    {
                        customInfo["AnnualFee1"] = drow["Annual Fee 1"].ToString();
                    }
                    if (isAnnualFee2Decimal == true)
                    {
                        customInfo["AnnualFee2"] = "$" + String.Format("{0:#,0.00}", drow["Annual Fee 2"]);
                    }
                    else
                    {
                        customInfo["AnnualFee2"] = String.Format("{0:#,0.00}", drow["Annual Fee 2"]);
                    }
                    if (isAnnualFee3Decimal == true)
                    {
                        customInfo["AnnualFee3"] = "$" + String.Format("{0:#,0.00}", drow["Annual Fee 3"]);
                    }
                    else
                    {
                        customInfo["AnnualFee3"] = drow["Annual Fee 3"].ToString();
                    }
                    if (isAnnualFee4Decimal == true)
                    {
                        customInfo["AnnualFee4"] = "$" + String.Format("{0:#,0.00}", drow["Annual Fee 4"]);
                    }
                    else
                    {
                        customInfo["AnnualFee4"] = drow["Annual Fee 4"].ToString();
                    }
                    if (isAnnualFee5Decimal == true)
                    {
                        customInfo["AnnualFee5"] = "$" + String.Format("{0:#,0.00}", drow["Annual Fee 5"]);
                    }
                    else
                    {
                        customInfo["AnnualFee5"] = drow["Annual Fee 5"].ToString();
                    }
                    if (isAnnualFee6Decimal == true)
                    {
                        customInfo["AnnualFee6"] = "$" + String.Format("{0:#,0.00}", drow["Annual Fee 6"]);
                    }
                    else
                    {
                        customInfo["AnnualFee6"] = drow["Annual Fee 6"].ToString();
                    }
                    if (isAnnualFee7Decimal == true)
                    {
                        customInfo["AnnualFee7"] = "$" + String.Format("{0:#,0.00}", drow["Annual Fee 7"]);
                    }
                    else
                    {
                        customInfo["AnnualFee7"] = drow["Annual Fee 7"].ToString();
                    }
                    if (isAnnualFee8Decimal == true)
                    {
                        customInfo["AnnualFee8"] = "$" + String.Format("{0:#,0.00}", drow["Annual Fee 8"]);
                    }
                    else
                    {
                        customInfo["AnnualFee8"] = drow["Annual Fee 8"].ToString();
                    }
                    if (isAnnualFee9Decimal == true)
                    {
                        customInfo["AnnualFee9"] = "$" + String.Format("{0:#,0.00}", drow["Annual Fee 9"]);
                    }
                    else
                    {
                        customInfo["AnnualFee9"] = drow["Annual Fee 9"].ToString();
                    }
                    if (isAnnualFee10Decimal == true)
                    {
                        customInfo["AnnualFee10"] = "$" + String.Format("{0:#,0.00}", drow["Annual Fee 10"]);
                    }
                    else
                    {
                        customInfo["AnnualFee10"] = drow["Annual Fee 10"].ToString();
                    }
                    if (isAnnualFee11Decimal == true)
                    {
                        customInfo["AnnualFee11"] = "$" + String.Format("{0:#,0.00}", drow["Annual Fee 11"]);
                    }
                    else
                    {
                        customInfo["AnnualFee11"] = drow["Annual Fee 11"].ToString();
                    }
                    if (isAnnualFee12Decimal == true)
                    {
                        customInfo["AnnualFee12"] = "$" + String.Format("{0:#,0.00}", drow["Annual Fee 12"]);
                    }
                    else
                    {
                        customInfo["AnnualFee12"] = String.Format("{0:#,0.00}", drow["Annual Fee 12"]);
                    }
                    if (isAnnualFee13Decimal == true)
                    {
                        customInfo["AnnualFee13"] = "$" + String.Format("{0:#,0.00}", drow["Annual Fee 13"]);
                    }
                    else
                    {
                        customInfo["AnnualFee13"] = drow["Annual Fee 13"].ToString();
                    }
                    if (isAnnualFee14Decimal == true)
                    {
                        customInfo["AnnualFee14"] = "$" + String.Format("{0:#,0.00}", drow["Annual Fee 14"]);
                    }
                    else
                    {
                        customInfo["AnnualFee14"] = drow["Annual Fee 14"].ToString();
                    }
                    if (isAnnualFee15Decimal == true)
                    {
                        customInfo["AnnualFee15"] = "$" + String.Format("{0:#,0.00}", drow["Annual Fee 15"]);
                    }
                    else
                    {
                        customInfo["AnnualFee15"] = drow["Annual Fee 15"].ToString();
                    }
                    if (isAnnualFee16Decimal == true)
                    {
                        customInfo["AnnualFee16"] = "$" + String.Format("{0:#,0.00}", drow["Annual Fee 16"]);
                    }
                    else
                    {
                        customInfo["AnnualFee16"] = drow["Annual Fee 16"].ToString();
                    }
                    if (isAnnualFee17Decimal == true)
                    {
                        customInfo["AnnualFee17"] = "$" + String.Format("{0:#,0.00}", drow["Annual Fee 17"]);
                    }
                    else
                    {
                        customInfo["AnnualFee17"] = drow["Annual Fee 17"].ToString();
                    }
                    if (isAnnualFee18Decimal == true)
                    {
                        customInfo["AnnualFee18"] = "$" + String.Format("{0:#,0.00}", drow["Annual Fee 18"]);
                    }
                    else
                    {
                        customInfo["AnnualFee18"] = drow["Annual Fee 18"].ToString();
                    }
                    if (isAnnualFee19Decimal == true)
                    {
                        customInfo["AnnualFee19"] = "$" + String.Format("{0:#,0.00}", drow["Annual Fee 19"]);
                    }
                    else
                    {
                        customInfo["AnnualFee19"] = drow["Annual Fee 19"].ToString();
                    }
                    if (isAnnualFee20Decimal == true)
                    {
                        customInfo["AnnualFee20"] = "$" + String.Format("{0:#,0.00}", drow["Annual Fee 20"]);
                    }
                    else
                    {
                        customInfo["AnnualFee20"] = drow["Annual Fee 20"].ToString();
                    }
                    customInfo["OptionalService1"] = drow["Optional Services/Features 1"].ToString();
                    customInfo["OptionalService2"] = drow["Optional Services/Features 2"].ToString();
                    customInfo["OptionalDescription1"] = drow["Optional Description 1"].ToString();
                    customInfo["OptionalDescription2"] = drow["Optional Description 2"].ToString();

                    var isOptionalImplementationFee1Decimal = Decimal.TryParse(drow["Optional Implmentation Fee 1"].ToString(), out OptionalImplementationFee1);
                    var isOptionalImplementationFee2Decimal = Decimal.TryParse(drow["Optional Implmentation Fee 2"].ToString(), out OptionalImplementationFee2);

                    if (isOptionalImplementationFee1Decimal == true)
                    {
                        customInfo["OptionalImplementationFee1"] = "$" + String.Format("{0:#,0.00}", drow["Optional Implmentation Fee 1"]);
                    }
                    else
                    {
                        customInfo["OptionalImplementationFee1"] = drow["Optional Implmentation Fee 1"].ToString();
                    }
                    if (isOptionalImplementationFee2Decimal == true)
                    {
                        customInfo["OptionalImplementationFee2"] = "$" + String.Format("{0:#,0.00}", drow["Optional Implmentation Fee 2"]);
                    }
                    else
                    {
                        customInfo["OptionalImplementationFee2"] = drow["Optional Implmentation Fee 2"].ToString();
                    }

                    var isOptionalAnnualFee1Decimal = Decimal.TryParse(drow["Optional Annual Fee 1"].ToString(), out OptionalAnnualFee1);
                    var isOptionalAnnualFee2Decimal = Decimal.TryParse(drow["Optional Annual Fee 2"].ToString(), out OptionalAnnualFee2);
                    if (isOptionalAnnualFee1Decimal == true)
                    {
                        customInfo["OptionalAnnualFee1"] = "$" + String.Format("{0:#,0.00}", drow["Optional Annual Fee 1"]);
                    }
                    else
                    {
                        customInfo["OptionalAnnualFee1"] = drow["Optional Annual Fee 1"].ToString();
                    }
                    if (isOptionalAnnualFee2Decimal == true)
                    {
                        customInfo["OptionalAnnualFee2"] = "$" + String.Format("{0:#,0.00}", drow["Optional Annual Fee 2"]);
                    }
                    else
                    {
                        customInfo["OptionalAnnualFee2"] = drow["Optional Annual Fee 2"].ToString();
                    }
                    customInfo["OptionalIncluded1"] = drow["Optional Included 1"].ToString();
                    customInfo["OptionalIncluded2"] = drow["Optional Included 2"].ToString();

                }

                return customInfo;

            }

            bool getDocType(DataTable dtInfo)
            {
                bool isBid = false;
                try
                {
                    foreach (DataRow drow in dtInfo.Rows)
                    {
                        isBid = drow["Template to use"].ToString().Equals("bids&tenders");
                    }
                }
                catch (Exception e)
                {
                    Log("Exception in GetDocType: " + e.Message);
                }

                return isBid;
            }

            string getFileName(DataTable dtInfo)
            {
                string FileName = "";
                try
                {
                    foreach (DataRow drow in dtInfo.Rows)
                    {
                        FileName = drow["Name of File"].ToString();
                    }
                }
                catch (Exception ex)
                {
                    Log("Exception in GetDocType: " + ex.Message);
                    strExceptionMessage = "Exception occured on generating the word document, this might be an issue with the supporting file, please check if the stockpile names are correct and please make sure that there is no merged cells." + "Error Code: " + ex.Message;
                }

                return FileName;
            }


            string CreateNewWordDoc(string newfolderName, bool IsBid, string fileName)
            {
                try
                {
                    //
                    string BidTemplate = bidTemplatePath;
                    string nonBidTemplate = nonBidTemplatePath;
                    Log("Start: " + MethodBase.GetCurrentMethod().Name);

                    //Create new file
                    string newFileName = fileName + Path.GetExtension(BidTemplate);
                    string newFilePath = Path.Combine(outputFolderPath, newfolderName, newFileName);
                    string newFolderPath = Path.GetDirectoryName(newFilePath);
                    strWordOutput = newFolderPath;
                    //string newNonBidFileName = fileName + "_" + DateTime.Today.ToString("ddMMyyyy") + Path.GetExtension(nonBidTemplate);
                    //string newNonBidFilePath = Path.Combine(outputFolderPath, newfolderName, newNonBidFileName);
                    //string newNonBidFolderPath = Path.GetDirectoryName(newNonBidFilePath);
                    //strWordOutput = newNonBidFolderPath;

                    Log("Create new file: " + newFileName);
                    Log("New Word Path: " + strWordOutput);

                    if (!Directory.Exists(newFolderPath))
                    {
                        Directory.CreateDirectory(newFolderPath);
                        Log("Created new directory: " + newFolderPath);
                    }
                    if (IsBid)
                    {
                        File.Copy(BidTemplate, newFilePath);
                    }
                    else
                    {
                        File.Copy(nonBidTemplate, newFilePath);
                    }


                    Log("Copied template to new file path");
                    Log("End: " + MethodBase.GetCurrentMethod().Name);
                    
                    return newFilePath;

                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            void Log(object text)
            {
                bool isLogCreated = false;
                StringBuilder sb = new StringBuilder();
                try
                {
                    string logFileName = "Logs_UpdateWordDoc";
                    string today = DateTime.Today.ToShortDateString().Replace("/", "_") + ".txt";
                    if (!isLogCreated)
                    {
                        if (System.IO.Directory.Exists(logDirectory))
                        {
                            if (!File.Exists(System.IO.Path.Combine(logDirectory, logFileName + today)))
                            {
                                FileStream fs = File.Create(System.IO.Path.Combine(logDirectory, logFileName + today));
                                fs.Close();
                            }
                        }
                        else
                        {
                            System.IO.Directory.CreateDirectory(logDirectory);
                            FileStream fs = File.Create(System.IO.Path.Combine(logDirectory, logFileName + today));
                            fs.Close();
                        }
                        isLogCreated = true;
                    }
                    //Have not used else because the first text won't get appended when file gets created
                    if (isLogCreated)
                    {
                        sb.Clear();
                        sb.Append(DateTime.Now.ToString() + " " + text.ToString());
                        File.AppendAllText(System.IO.Path.Combine(logDirectory, logFileName + today), "\r\n" + sb.ToString());
                    }
                }
                catch (Exception ex)
                {
                    Log("Exception occurred on Log Method:" + ex.ToString());
                }
            }

              
           // Outputs
           return (ctx) => {
                ExceptionMessage.Set(ctx, strExceptionMessage);
                WordOutputPath.Set(ctx, strWordOutput);
           };
        }

        #endregion
    }
}

