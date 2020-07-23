using Microsoft.SharePoint.Client;
using System;
using System.Data;
using System.Linq;
using System.IO;
using System.Net;
using PDF_Report_Generator.DataAccess;
using System.Text;
using System.Drawing;
using IronPdf;
using PDF_Report_Generator.Logging;

namespace PDF_Report_Generator
{
    class CustomerWeeklyReport
    {
        private DataTable dtDetails = new DataTable();
        private static readonly string[] _validImageExtensions = { ".jpg", ".bmp", ".gif", ".png", ".jpeg" };

        public static void CreateReport(string spUserName, string spUserPWD, string spListID, HtmlToPdf customerWeekly_html, HtmlToPdf customerWeekly_html_coverpage, string isDebugMode)
        {
            try
            {
                LogFile_CustomerWeekly.WriteBlankLine();
                LogFile_CustomerWeekly.WriteLogMessage("=================================================================================================++++++++++");
                LogFile_CustomerWeekly.WriteLogMessage("Starting PDF Generator for Customer Weekly Report for List ID: " + spListID.ToString());
               
                //Get the list of records to process
                DataTable dt = GetDetailsList(spListID);

                if (dt.Rows.Count < 1)
                {
                    string errMsg = "ERROR CreateReport: Unable to retrieve record from the DB for List ID: " + spListID.ToString();
                    LogFile_CustomerWeekly.WriteLogMessage(errMsg);
                    PDF_Report_Generator.SaveErrorToDB(spListID,"LIST_ID", errMsg,1, "CustomerWeeklyReport");
                }

                foreach (DataRow row in dt.Rows)
                {
                    int listItemID = int.Parse(row["list_id"].ToString());
                    string jobNumber = row["job_number"].ToString();
                    string jobGUID = row["jobGUID"].ToString();

                    LogFile_CustomerWeekly.WriteLogMessage("Sharepoint ListID: " + listItemID.ToString());

                    String siteUrl = "https://mySharePointLists.com/reports";
                    String listName = "Customer Weekly Report";
                    NetworkCredential credentials = new NetworkCredential(spUserName, spUserPWD, "test");

                    using (ClientContext clientContext = new ClientContext(siteUrl))
                    {
                        LogFile_CustomerWeekly.WriteLogMessage("Started Attachment Download " + siteUrl);
                        clientContext.Credentials = credentials;

                        //Get the Site Collection
                        Site oSite = clientContext.Site;
                        clientContext.Load(oSite);
                        clientContext.ExecuteQuery();

                        // Get the Web
                        Web oWeb = clientContext.Web;
                        clientContext.Load(oWeb);
                        clientContext.ExecuteQuery();

                        CamlQuery query = new CamlQuery();
                        query.ViewXml = "<View><Query><Where><Geq><FieldRef Name='ID'/>" +
                                        "<Value Type='Number'>" + listItemID + "</Value></Geq></Where></Query><RowLimit>100</RowLimit></View>";

                        List oList = clientContext.Web.Lists.GetByTitle(listName);
                        clientContext.Load(oList);
                        clientContext.ExecuteQuery();

                        ListItemCollection items = oList.GetItems(query);
                        clientContext.Load(items);
                        clientContext.ExecuteQuery();
                                              

                       //create the header for each page
                        WebClient wc_pdf_header = new WebClient();
                        string pdf_header = wc_pdf_header.DownloadString("http://mySiteAssest.com/HTML_Report_Templates/CustomerWeeklyReport/PDF_Header_Template.html");
                        StringBuilder sb_header = new StringBuilder(pdf_header);
                        byte[] bannerImgAsByteArray;
                        using (var webClient = new WebClient())
                        {
                            bannerImgAsByteArray = webClient.DownloadData("http://mySiteAssest.com/Images/CreatePDFReport_banner.png");
                        }
                        //replace the banner image placeholder                        
                        sb_header.Replace("{banner_image}", photoURI(bannerImgAsByteArray));

                        //add the header html to all body pages
                        customerWeekly_html.PrintOptions.Header = new HtmlHeaderFooter()
                        {
                            HtmlFragment = sb_header.ToString(),
                            Height = 35
                        };

                        //create the footer for each page
                        WebClient wc_pdf_footer = new WebClient();
                        string pdf_footer = wc_pdf_footer.DownloadString("http://mySiteAssest.com/HTML_Report_Templates/CustomerWeeklyReport/PDF_Footer_Template.html");

                        //replace the field placeholders
                        StringBuilder sb_footer = new StringBuilder(pdf_footer);
                        sb_footer.Replace("{job_number}", row["job_number"].ToString());
                        sb_footer.Replace("{project_name}", row["project_name"].ToString());
                        sb_footer.Replace("{location}", row["location"].ToString());

                        //add the footer html to all body pages
                        customerWeekly_html.PrintOptions.Footer = new HtmlHeaderFooter()
                        {
                            HtmlFragment = sb_footer.ToString(),
                            Height = 32
                        };

                        //Get the pdf coverpage template
                        WebClient wc_pdf_coverpage = new WebClient();
                        string pdf_coverpage = wc_pdf_coverpage.DownloadString("http://mySiteAssest.com/HTML_Report_Templates/CustomerWeeklyReport/PDF_CoverPage_Template.html");

                        //create the background image for the cover page
                        byte[] bgImgAsByteArray;
                        using (var webClient = new WebClient())
                        {
                            bgImgAsByteArray = webClient.DownloadData("http://mySiteAssest.com/Images/CreatePDFReport_Background_Fade.png");
                        }

                        StringBuilder sb_coverpage = new StringBuilder(pdf_coverpage);
                        sb_coverpage.Replace("{background_image}", photoURI(bgImgAsByteArray));
                        sb_coverpage.Replace("{cp_job_number}", row["job_number"].ToString());
                        sb_coverpage.Replace("{cp_project_name}", row["project_name"].ToString());
                        sb_coverpage.Replace("{cp_location}", row["location"].ToString());
                        sb_coverpage.Replace("{report_date}", row["week_ending"].ToString());

                        //add the header html to the coverpage
                        customerWeekly_html_coverpage.PrintOptions.Header = new HtmlHeaderFooter()
                        {
                            HtmlFragment = sb_header.ToString(),
                            Height = 35
                        };

                        //replace the coverpage footer placeholders - this page will not have any data in the footer
                        StringBuilder sb_footer_coverpage = new StringBuilder(pdf_footer);
                        sb_footer_coverpage.Replace("Job#: </strong>{job_number}", "&nbsp;");
                        sb_footer_coverpage.Replace("Project Name: </strong>{project_name}", "&nbsp;");
                        sb_footer_coverpage.Replace("Location: </strong>{location}", "&nbsp;");

                        //add the footer html to the coverpage
                        customerWeekly_html_coverpage.PrintOptions.Footer = new HtmlHeaderFooter()
                        {
                            HtmlFragment = sb_footer_coverpage.ToString(),
                            Height = 32
                        };

                        //Get the pdf body template
                        WebClient wc_pdf_Body = new WebClient();
                        string pdf_body = wc_pdf_Body.DownloadString("http://mySiteAssest.com/HTML_Report_Templates/CustomerWeeklyReport/PDF_Body_Template.html");

                        StringBuilder sb_body = new StringBuilder(pdf_body);

                        DataTable dtVP = GeNARContacts();
                        bool narContact1Flag = false;
                        bool narContact2Flag = false;
                        bool narContact3Flag = false;

                        foreach(DataRow narContact in dtVP.Rows)
                        {
                            if (narContact["OrderBy"].ToString() == "1")
                            {
                                sb_body.Replace("{vp_name}", narContact["FullName"].ToString());
                                sb_body.Replace("{vp_title}", narContact["PositionIdName"].ToString());
                                sb_body.Replace("{vp_phone}", narContact["address1_telephone1"].ToString());
                                sb_body.Replace("{vp_email}", narContact["internalemailaddress"].ToString());
                                narContact1Flag = true;
                            }
                            if (narContact["OrderBy"].ToString() == "2")
                            {
                                sb_body.Replace("{sdr_name}", narContact["FullName"].ToString());
                                sb_body.Replace("{sdr_title}", narContact["PositionIdName"].ToString());
                                sb_body.Replace("{sdr_phone}", narContact["address1_telephone1"].ToString());
                                sb_body.Replace("{sdr_email}", narContact["internalemailaddress"].ToString());
                                narContact2Flag = true;
                            }
                            if (narContact["OrderBy"].ToString() == "3")
                            {
                                sb_body.Replace("{mdr_name}", narContact["FullName"].ToString());
                                sb_body.Replace("{mdr_title}", narContact["PositionIdName"].ToString());
                                sb_body.Replace("{mdr_phone}", narContact["address1_telephone1"].ToString());
                                sb_body.Replace("{mdr_email}", narContact["internalemailaddress"].ToString());
                                narContact3Flag = true;
                            }
                        }

                        //clear out placeholders for the NAR contacts that we do not have
                        if (!narContact1Flag)
                        {
                            sb_body.Replace("{vp_name}", "");
                            sb_body.Replace("{vp_title}", "");
                            sb_body.Replace("{vp_phone}", "");
                            sb_body.Replace("{vp_email}", "");
                        }

                        if (!narContact2Flag)
                        {
                            sb_body.Replace("{sdr_name}", "");
                            sb_body.Replace("{sdr_title}", "");
                            sb_body.Replace("{sdr_phone}", "");
                            sb_body.Replace("{sdr_email}", "");
                        }

                        if (!narContact3Flag)
                        {
                            sb_body.Replace("{mdr_name}", "");
                            sb_body.Replace("{mdr_title}", "");
                            sb_body.Replace("{mdr_phone}", "");
                            sb_body.Replace("{mdr_email}", "");
                        }

                        //replace the field placeholders in the body
                        sb_body.Replace("{pm_name}", row["pm_name"].ToString());
                        sb_body.Replace("{pm_phone}", row["pm_phone"].ToString());
                        sb_body.Replace("{pm_email}", row["pm_email"].ToString());
                        sb_body.Replace("{apm_name}", row["apm_name"].ToString());
                        sb_body.Replace("{apm_phone}", row["apm_phone"].ToString());
                        sb_body.Replace("{apm_email}", row["apm_email"].ToString());
                        sb_body.Replace("{job_type}", row["job_type"].ToString());
                        sb_body.Replace("{job_size}", row["job_size"].ToString());
                        sb_body.Replace("{est_completion_date}", row["est_completion_date"].ToString());
                        sb_body.Replace("{total_squares}", row["total_squares"].ToString());
                        sb_body.Replace("{percent_complete}", row["percent_complete"].ToString());
                        sb_body.Replace("{squares_installed}", row["squares_installed"].ToString());
                        sb_body.Replace("{total_squares}", row["total_squares"].ToString());


                        //Create the invdividual photo template
                        WebClient wc_pdf_photopage = new WebClient();
                        string photopage_Template = wc_pdf_photopage.DownloadString("http://mySiteAssest.com/HTML_Report_Templates/CustomerWeeklyReport/PDF_PhotosPage_Template.html");

                        StringBuilder sp_photopage = new StringBuilder(photopage_Template);
                        

                        //Create the invdividual photo template
                        WebClient wc_pdf_photo = new WebClient();
                        string photo_Template = wc_pdf_photo.DownloadString("http://mySiteAssest.com/HTML_Report_Templates/CustomerWeeklyReport/PDF_Photo_Template.html");

                        StringBuilder photoDivs = new StringBuilder();
                        bool hasImages = false;

                        foreach (ListItem listItem in items)
                        {
                            if (Int32.Parse(listItem["ID"].ToString()) == listItemID)
                            {

                                string folderURL = oSite.Url + "/customerWeeklyRpts/Lists/" + listName + "/Attachments/" + listItem["ID"];
                                Folder folder = oWeb.GetFolderByServerRelativeUrl(folderURL);

                                clientContext.Load(folder);

                                try
                                {
                                    clientContext.ExecuteQuery();
                                }
                                catch (ServerException ex)
                                {
                                    LogFile_CustomerWeekly.WriteLogMessage(ex.Message);
                                    LogFile_CustomerWeekly.WriteLogMessage("No Attachment for ID " + listItem["ID"].ToString());
                                    LogFile_CustomerWeekly.WriteBlankLine();
                                    LogFile_CustomerWeekly.WriteBlankLine();
                                }

                                FileCollection attachments = folder.Files;
                                clientContext.Load(attachments);
                                clientContext.ExecuteQuery();

                                //Set the counts used to insert rows of images - 3 images per row
                                int imageCnt = 0;
                                int totalImgCnt = 0;

                                string photo1 = "";
                                string photo2 = "";
                                

                                //Loop through the photos attached to the sharepoint report
                                foreach (Microsoft.SharePoint.Client.File oFile in folder.Files)
                                {
                                    FileInfo fiPhoto = new FileInfo(oFile.Name);
                                    WebClient client1 = new WebClient();
                                    client1.Credentials = credentials;

                                    LogFile_CustomerWeekly.WriteLogMessage("Downloading " + oFile.ServerRelativeUrl);
                                    if (IsImageExtension(fiPhoto.Extension))
                                    {
                                        //convert the photo to a byte image and then resize it so that they are all uniform.
                                        byte[] fileContents = client1.DownloadData("https://mySharePointLists.com" + oFile.ServerRelativeUrl);
                                        byte[] resizedImage = CreateThumbnail(fileContents, 300);

                                        //Get the extension of the file - we can only use images on the PDF - jpg, jpeg, png, gif, bmp
                                        LogFile_CustomerWeekly.WriteLogMessage("Photo Extension " + fiPhoto.Extension);

                                        imageCnt++;
                                        totalImgCnt++;
                                        hasImages = true;
                                        //Add the image to the photo area template
                                        if (imageCnt == 2)
                                        {
                                            photoDivs.Append(photo_Template);
                                            //now convert the byte image to a DataUri stream so we can include it in the pdf
                                            photo1 = photoURI(resizedImage);
                                            photoDivs.Replace("{photo_number}", totalImgCnt.ToString());
                                            photoDivs.Replace("{photo_source}", photo1);
                                            photoDivs.Append("</tr>");

                                            //Reset our image variable for the next row
                                            imageCnt = 0;
                                        }
                                        else
                                        {
                                            photoDivs.Append("<tr>");
                                            photoDivs.Append(photo_Template);
                                            //now convert the byte image to a DataUri stream so we can include it in the pdf
                                            photo2 = photoURI(resizedImage);
                                            photoDivs.Replace("{photo_number}", totalImgCnt.ToString());
                                            photoDivs.Replace("{photo_source}", photo2);
                                        }

                                    }
                                    else
                                    {
                                        LogFile_CustomerWeekly.WriteLogMessage("Invalid Photo Extension " + fiPhoto.Extension);
                                    }
                                }

                                //Add the last row of images
                                if (imageCnt < 2 && hasImages == true)
                                {
                                    photoDivs.Append("</tr>");
                                }
                            }
                        }
                        if (!hasImages)
                        {
                            //There was no valid image for this job so leave the photo are blank
                            sb_body.Replace("{photos_page}", "");
                        }
                        else
                        {
                            //There was at least one valid image for this job so insert the photopage template into the body
                            sb_body.Replace("{photos_page}",sp_photopage.ToString());
                            //Then insert the photo template into the body template
                            sb_body.Replace("{photos}", photoDivs.ToString());
                        }


                        //Save the completed PDF document to the sharepoint doc location for the job
                        bool spSave = PDF_Report_Generator.createCustomerWeeklyReport(sb_coverpage, customerWeekly_html_coverpage, sb_body, customerWeekly_html, jobNumber, jobGUID,spListID,"1",isDebugMode);
                        var pdfFileName = "CustomerWeeklyReport_" + jobNumber + "_" + spListID.ToString() + "_" + DateTime.Now.ToString("yyyyMMdd") + ".pdf";
                        
                        if (spSave)
                        {
                            LogFile_CustomerWeekly.WriteLogMessage("SUCCESS: File " + pdfFileName + " saved to sharepoint doc location for job #: " + jobNumber);
                        }
                        else
                        {
                            LogFile_CustomerWeekly.WriteLogMessage("ERROR: File " + pdfFileName + " was not saved to sharepoint doc location for job: " + jobNumber);
                        }

                        //for debugging only -we will not save the pdf to a file location
                        //PDFDoc_Report.SaveAs(@"C:\Projects\Projects\Create_PDF_Reports\Reports\" + pdfFileName);

                        LogFile_CustomerWeekly.WriteLogMessage("=================================================================================================++++++++++");
                        LogFile_CustomerWeekly.WriteBlankLine();
                    }
                }
            }
            catch (Exception e)
            {
                LogFile_CustomerWeekly.WriteLogMessage("ERROR CreateReport: Unable to create PDF report - " + e.Message);
                LogFile_CustomerWeekly.WriteBlankLine();
                LogFile_CustomerWeekly.WriteBlankLine();

                string errMsg = "ERROR CreateReport: Unable to create PDF report - " + e.Message;
                LogFile_CustomerWeekly.WriteLogMessage(errMsg);
                PDF_Report_Generator.SaveErrorToDB(spListID,"LIST_ID", errMsg,1, "CustomerWeeklyReport");
            }
        }

        public static DataTable GetDetailsList(string listid)
        {
            DataQueries dataAccess = new DataQueries();

            string query = "SELECT TOP 1 [list_id] ,[jobno] AS [job_number],[job_name] AS [project_name],[est_completion_date],[job_size] ,[job_type] ,[job_city] ,[job_state],[job_city] + ', ' +[job_state] AS [location]  " +
                            " ,[NAR_JobProjectManagerId],[pm_name],[pm_email],[pm_phone],[pm_mobile],[nar_jobasstprojectmanagerid],[apm_name],[apm_email] " +
                            " ,[apm_phone],[apm_mobile],[total_squares],[percent_complete],[squares_installed],[week_ending],[ModifiedOn],[jobGUID] " +
                            " ,[job_type], [job_size] " +
                             " FROM [PDF_RPT_DATA].[dbo].[View_Customer_Weekly_Report_Data] " +
                            " WHERE [list_id] = " + listid +
                            " ORDER BY [ModifiedOn] DESC";

            DataTable dt = dataAccess.ReadDataTable(query,"CRM");

            return dt;
        }

        public static DataTable GeNARContacts()
        {
            DataQueries dataAccess = new DataQueries();

            string query = "SELECT FullName,title,IsDisabled,PositionId,PositionIdName,address1_telephone1,internalemailaddress,OrderBy " +
                            "FROM [PDF_RPT_DATA].[dbo].[View_CustomerWeekly_Contacts]";
                            
            DataTable dt = dataAccess.ReadDataTable(query, "CRM");

            return dt;
        }

        public static bool IsImageExtension(string ext)
        {
            return _validImageExtensions.Contains(ext.ToLower());
        }

        public static string photoURI(byte[] resizedImage)
        {
            //convert a byte image to a DataURI stream
            Image photoImage = (Bitmap)((new ImageConverter()).ConvertFrom(resizedImage));
            var DataURI = IronPdf.Util.ImageToDataUri(photoImage);
            return DataURI;
        }

        // (RESIZE an image in a byte[] variable. Will keep the aspect ratio)  
        public static byte[] CreateThumbnail(byte[] PassedImage, int LargestSide)
        {
            byte[] ReturnedThumbnail;

            using (MemoryStream StartMemoryStream = new MemoryStream(), NewMemoryStream = new MemoryStream())
            {
                // write the string to the stream  
                StartMemoryStream.Write(PassedImage, 0, PassedImage.Length);

                // create the start Bitmap from the MemoryStream that contains the image  
                Bitmap startBitmap = new Bitmap(StartMemoryStream);

                // set thumbnail height and width proportional to the original image.  
                int newHeight;
                int newWidth;
                double HW_ratio;
                if (startBitmap.Height > startBitmap.Width)
                {
                    newHeight = LargestSide;
                    HW_ratio = (double)((double)LargestSide / (double)startBitmap.Height);
                    newWidth = (int)(HW_ratio * (double)startBitmap.Width);
                }
                else
                {
                    newWidth = LargestSide;
                    HW_ratio = (double)((double)LargestSide / (double)startBitmap.Width);
                    newHeight = (int)(HW_ratio * (double)startBitmap.Height);
                }

                // create a new Bitmap with dimensions for the thumbnail.  
                Bitmap newBitmap = new Bitmap(newWidth, newHeight);

                // Copy the image from the START Bitmap into the NEW Bitmap.  
                // This will create a thumnail size of the same image.  
                newBitmap = ResizeImage(startBitmap, newWidth, newHeight);

                // Save this image to the specified stream in the specified format.  
                newBitmap.Save(NewMemoryStream, System.Drawing.Imaging.ImageFormat.Jpeg);

                // Fill the byte[] for the thumbnail from the new MemoryStream.  
                ReturnedThumbnail = NewMemoryStream.ToArray();
            }

            // return the resized image as a string of bytes.  
            return ReturnedThumbnail;
        }

        // Resize a Bitmap  called from CreateThumbnail()
        private static Bitmap ResizeImage(Bitmap image, int width, int height)
        {
            Bitmap resizedImage = new Bitmap(width, height);
            using (Graphics gfx = Graphics.FromImage(resizedImage))
            {
                gfx.DrawImage(image, new Rectangle(0, 0, width, height),
                    new Rectangle(0, 0, image.Width, image.Height), GraphicsUnit.Pixel);
            }
            return resizedImage;
        }

        //Save the PdfDocument stream to sharepoint
        public static bool SaveToSharepointFolder(PdfDocument PDF, string jobGUID, string username, string pwd,string filename,string spListID)
        {
            bool results = false;
            try
            {
                LogFile_CustomerWeekly.WriteLogMessage("Saving PDF to Sharepoint");

                String sharePointSite = "https://mySharePointLists.com/reports/CustomerReports/";
                String sharePointRootSite = "https://mySharePointLists.com/";
                string listname = "jobs";
                string folderName = "Customer_Weekly_Reports";
                String folderSiteURL = "/" + jobGUID + "/" + folderName + "/" + filename;
                string folderServerRelativeUrl = "/jobs/"+jobGUID+ "/Customer_Weekly_Reports/";
                string folderPath = jobGUID + "/Customer_Weekly_Reports";
                string savePDFURL = "/CustomerReports/jobs/" + jobGUID + "/Customer_Weekly_Reports/" + filename;

                NetworkCredential spCredentials = new NetworkCredential(username, pwd, "test");

                //convert the pdf doc to binary and then to a memory stream
                byte[] data = PDF.BinaryData;
                MemoryStream msPDF = new MemoryStream(data);
                //make sure the the memory stream is at the beginning
                msPDF.Seek(0, SeekOrigin.Begin);

                using (ClientContext ctx = new ClientContext(sharePointRootSite))
                {
                    ctx.Credentials = spCredentials;

                    //Get the Site Collection
                    Site oSite = ctx.Site;
                    ctx.Load(oSite);
                    ctx.ExecuteQuery();

                    //Create the folder in the jobs sharepoint doc location for the report if it does not exist yet
                    bool savePDF = CreateFolder(sharePointSite, listname, jobGUID,folderName, spCredentials,spListID);

                    if (savePDF)
                    {
                        using (ClientContext ctxSavePDF = new ClientContext(sharePointRootSite))
                        {
                            ctxSavePDF.Credentials = spCredentials;
                            Microsoft.SharePoint.Client.File.SaveBinaryDirect(ctxSavePDF, savePDFURL, msPDF, true);
                            results = true;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                LogFile_CustomerWeekly.WriteLogMessage("ERROR SaveToSharepointFolder: Could not PDF save to SharePoint - " + e.Message);
                LogFile_CustomerWeekly.WriteBlankLine();
                LogFile_CustomerWeekly.WriteBlankLine();

                string errMsg = "ERROR SaveToSharepointFolder: Could not save PDF to SharePoint - " + e.Message;
                LogFile_CustomerWeekly.WriteLogMessage(errMsg);
                PDF_Report_Generator.SaveErrorToDB(spListID, "LIST_ID", errMsg,1, filename);
            }

            return results;

        }

        private static bool CreateFolder(string siteUrl, string listName, string relativePath, string folderName, NetworkCredential spCredentials,string spListID)
        {
            bool result = false;

            try
            {
                using (ClientContext clientContext = new ClientContext(siteUrl))
                {
                    clientContext.Credentials = spCredentials;

                    Web web = clientContext.Web;
                    List list = web.Lists.GetByTitle(listName);

                    ListItemCreationInformation newItem = new ListItemCreationInformation();
                    newItem.UnderlyingObjectType = FileSystemObjectType.Folder;
                    newItem.FolderUrl = siteUrl + listName;
                    if (!relativePath.Equals(string.Empty))
                    {
                        newItem.FolderUrl += "/" + relativePath;
                    }
                    newItem.LeafName = folderName;
                    ListItem item = list.AddItem(newItem);
                    item.Update();
                    clientContext.ExecuteQuery();

                    result = true;
                }
            }
            catch (Exception Ex)
            {
                string msg = Ex.Message.ToString();
                //check the error message. If it indicates that the folder already exists then send back true 
                if (msg.Contains("Customer_Weekly_Reports already exists"))
                {
                    result = true;
                }
                else
                {
                    LogFile_CustomerWeekly.WriteLogMessage("ERROR CreateFolder: Could not create SP Folder - " + msg);
                    //a different error occurred so send back false
                    result = false;

                    string errMsg = "ERROR CreateFolder: Could not create SP Folder -  " + msg;
                    LogFile_CustomerWeekly.WriteLogMessage(errMsg);
                    PDF_Report_Generator.SaveErrorToDB(spListID, "LIST_ID", errMsg,1, "CustomerWeeklyReport");
                }

            }

            return result;
        }


    }
}