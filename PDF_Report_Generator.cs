using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IronPdf;
using PDF_Report_Generator.Logging;
using PDF_Report_Generator.DataAccess;

namespace PDF_Report_Generator
{
    class PDF_Report_Generator
    {
        public static string spUserName="";
        public static string spUserPWD = "";
        public static string spRptID = "";
        public static string spListID = "";
        public static string isDebugMode = "";

        static void Main(string[] args)
        {
            try
            {
               
                if (args == null || args.Length != 5)
                {
                    LogFile_PDFGenerator.WriteLogMessage("Missing user name, password, report name, list id, and/or debug flaf - could not generate PDF.");
                    LogFile_PDFGenerator.WriteLogMessage("Prompting for user name, password, report name, list ID, and debug flag.");
                    LogFile_PDFGenerator.WriteBlankLine();
                    LogFile_PDFGenerator.WriteBlankLine();

                    Console.WriteLine("Enter user name:");
                    spUserName = Console.ReadLine();
                    
                    Console.WriteLine("Enter user password:");
                    spUserPWD = Console.ReadLine();

                    Console.WriteLine("Select the report to generate:");
                    Console.WriteLine("   1 = Customer Weekly Report (require sharepoint list)");
                    spRptID = Console.ReadLine();
                                        
                    Console.WriteLine("Enter SharePoint list ID (enter 0 if there is not a sharepoint list):");
                    spListID = Console.ReadLine();

                    Console.WriteLine("Are you running this in test/debug mode? (enter Y or N):");
                    isDebugMode = Console.ReadLine();
                }
                else
                {
                    LogFile_PDFGenerator.WriteLogMessage("Parameters Passed: Count - " + args.Length.ToString());
                    spUserName = args[0].ToString();
                    spUserPWD = args[1].ToString();
                    spRptID = args[2].ToString();
                    spListID = args[3].ToString();
                    isDebugMode = args[4].ToString();

                    LogFile_PDFGenerator.WriteLogMessage("Parameters Passed: User - " + spUserName.ToString() + ",PWD: xxxxxxxx , RptID: " + spRptID.ToString() + " , List ID: " + spListID.ToString() + " , DebugMode: " + isDebugMode);
                }

                
                if(spRptID == "1")
                {
                    LogFile_PDFGenerator.WriteBlankLine();
                    LogFile_PDFGenerator.WriteLogMessage("=================================================================================================++++++++++");
                    LogFile_PDFGenerator.WriteLogMessage("Generate Customer Weekly Report PDF");

                    // Create a PDF from an HTML Template using IronPDF
                    HtmlToPdf customerWeekly_html = new HtmlToPdf();
                    customerWeekly_html.PrintOptions.MarginBottom = 0;
                    customerWeekly_html.PrintOptions.MarginTop = 0;
                    customerWeekly_html.PrintOptions.MarginLeft = 0;
                    customerWeekly_html.PrintOptions.MarginRight = 0;
                    customerWeekly_html.PrintOptions.PaperSize = PdfPrintOptions.PdfPaperSize.Letter;


                    // Create a PDF cover page from an HTML Template using IronPDF
                    HtmlToPdf customerWeekly_html_coverpage = new HtmlToPdf();
                    customerWeekly_html_coverpage.PrintOptions.MarginBottom = 0;
                    customerWeekly_html_coverpage.PrintOptions.MarginTop = 0;
                    customerWeekly_html_coverpage.PrintOptions.MarginLeft = 0;
                    customerWeekly_html_coverpage.PrintOptions.MarginRight = 0;
                    customerWeekly_html_coverpage.PrintOptions.PaperSize = PdfPrintOptions.PdfPaperSize.Letter;

                    CustomerWeeklyReport.CreateReport(spUserName,spUserPWD,spListID, customerWeekly_html, customerWeekly_html_coverpage, isDebugMode);
                }
            }
            catch (Exception e)
            {
                LogFile_PDFGenerator.WriteLogMessage("Error Messge: " + e.Message);
                LogFile_PDFGenerator.WriteBlankLine();
                LogFile_PDFGenerator.WriteBlankLine();

                string errMsg = "ERROR Main:  " + e.Message;
                LogFile_PDFGenerator.WriteLogMessage(errMsg);
                SaveErrorToDB(spListID,"LIST_ID", errMsg, Convert.ToInt32(spRptID), "-");
                LogFile_PDFGenerator.WriteLogMessage("=================================================================================================++++++++++");
                LogFile_PDFGenerator.WriteBlankLine();

            }
        }
                
        public static PdfDocument createPDFDoucment(StringBuilder html, HtmlToPdf page)
        {
            PdfDocument PDFDoc = page.RenderHtmlAsPdf(html.ToString());

            return PDFDoc;
         }

        public static PdfDocument mergePDFDoucments(PdfDocument page1, PdfDocument page2)
        {
            PdfDocument PDFDoc = PdfDocument.Merge(page1, page2);
            return PDFDoc;


        }

        public static void SaveErrorToDB(string srcID,string srcType, string errMsg,int spRptID, string rptName)
        {
            DataQueries dataAccess = new DataQueries();

            string query = "INSERT INTO [dbo].[PDF_GENERATOR_ERROR_LOG] " +
                            "([REPORT_ID],[PDF_REPORT_NAME],[ERROR_MESSAGE],[PDF_SOURCE_ID],[PDF_SOURCE_ID_TYPE],[DATE_CREATED]) " +
                            "VALUES " +
                            "("+ spRptID + ",'" + rptName + "','" + errMsg + "','"+ srcID +"','"+ srcType +"', GETDATE())";

            dataAccess.QryCommand(query, "PDF");
        }

        public static bool createCustomerWeeklyReport(StringBuilder sb_coverpage, HtmlToPdf customerWeekly_html_coverpage, StringBuilder sb_body, HtmlToPdf customerWeekly_html, string jobNumber, string jobGUID,string spListID,string spRptID, string isDebugMode)
        {
            //Save the completed PDF document to the sharepoint doc location for the job
            bool spSave = false;
            try
            {
                PdfDocument PDFDoc_coverpage = createPDFDoucment(sb_coverpage, customerWeekly_html_coverpage);
                PdfDocument PDFDoc_body = createPDFDoucment(sb_body, customerWeekly_html);
                PdfDocument PDFDoc_Report = mergePDFDoucments(PDFDoc_coverpage, PDFDoc_body);

                var pdfFileName = "CustomerWeeklyReport_" + jobNumber + "_" + spListID.ToString() + "_" + DateTime.Now.ToString("yyyyMMdd") + ".pdf";
                if (isDebugMode == "N")
                {
                    spSave = CustomerWeeklyReport.SaveToSharepointFolder(PDFDoc_Report, jobGUID, spUserName, spUserPWD, pdfFileName, spListID);
                }
                else
                {
                    //for debugging only -we will not save the pdf to a file location
                    PDFDoc_Report.SaveAs(@"C:\Projects\Projects\Create_PDF_Reports\Reports\CustomerWeekly\" + pdfFileName);
                }

               

            }
            catch (Exception e)
            {
                LogFile_PDFGenerator.WriteLogMessage("ERROR createCustomerWeeklyReport: " + e.Message);
                LogFile_PDFGenerator.WriteBlankLine();
                LogFile_PDFGenerator.WriteBlankLine();

                LogFile_CustomerWeekly.WriteLogMessage("ERROR createCustomerWeeklyReport: " + e.Message);
                LogFile_CustomerWeekly.WriteBlankLine();
                LogFile_CustomerWeekly.WriteBlankLine();

                string errMsg = "ERROR createCustomerWeeklyReport:  " + e.Message;
                string rptName = "CustomerWeeklyReport_" + jobNumber + "_" + spListID;
                LogFile_CustomerWeekly.WriteLogMessage(errMsg);
                SaveErrorToDB(spListID, "LIST_ID", errMsg, Convert.ToInt32(spRptID), rptName);

                LogFile_CustomerWeekly.WriteLogMessage("=================================================================================================++++++++++");
                LogFile_CustomerWeekly.WriteBlankLine();
            }

            return spSave;
        }
    }
}
