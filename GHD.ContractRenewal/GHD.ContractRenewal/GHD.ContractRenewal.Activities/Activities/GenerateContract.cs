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

            //Output
            string strWordOutput = "";
            string strExceptionMessage = "";
            
            


            try
            {
                UpdateWordDoc(dataTableInfo, dataTableAnnual, dataTableOptional);
            }
            catch(Exception e)
            {

            }
           
            void UpdateWordDoc(DataTable dtInfo, DataTable dtAnnual, DataTable dtOptional)
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
                     */
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
                    customInfo["Delivery Timing"] = drow["Estimated Delivery Timing"].ToString();
                    customInfo["License Term"] = drow["License Term"].ToString();
                    customInfo["Payment Terms"] = drow["Payment Terms"].ToString();
                    customInfo["Expiry Date"] = drow["Quote Expiry Date"].ToString();
                    customInfo["Executive"] = drow["Account Executive"].ToString();
                    customInfo["Executive Email"] = drow["Account Executive Email"].ToString();
                    customInfo["Executive Phone"] = drow["Account Executive Phone"].ToString();
                    customInfo["Purpose"] = drow["Purpose"].ToString();
                    customInfo["Total Implementation"] = drow["Total Implmentation Fee"].ToString();
                    customInfo["Total Annual"] = drow["Total Annual Fee"].ToString();
                    customInfo["Acceptance Criteria"] = drow["Acceptance Criteria"].ToString();
                    customInfo["Delivery Schedule"] = drow["Delivery Schedule"].ToString();
                    customInfo["Exclusions"] = drow["Exclusions and Assumptions"].ToString();
                    customInfo["Payment Schedule"] = drow["Payment Schedule On-Time"].ToString();
                    customInfo["Year Fee"] = drow["Payment Schedule Year's fees"].ToString();
                    customInfo["Sign Name"] = drow["Customer Sign Print Name"].ToString();
                    customInfo["Sign Title"] = drow["Customer Sign Print Title"].ToString();
                    customInfo["Sign Date"] = drow["Customer Sign Date"].ToString();
                    customInfo["GHD Sign Date"] = drow["GHD Digital Sign Date"].ToString();
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
                    string newFileName = fileName + "_" + DateTime.Today.ToString("ddMMyyyy") + Path.GetExtension(BidTemplate);
                    string newFilePath = Path.Combine(outputFolderPath, newfolderName, newFileName);
                    string newFolderPath = Path.GetDirectoryName(newFilePath);
                    strWordOutput = newFolderPath;


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

