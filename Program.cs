using System;
using System.ComponentModel;
using System.Configuration;
using System.Net.Mail;
using System.Net.NetworkInformation;

using databaseAPI;

using EASendMail; //add EASendMail namespace (This needs the license code)

using GNAgeneraltools;

using GNAspreadsheettools;

using OfficeOpenXml;

using SmtpClient = EASendMail.SmtpClient;

namespace DisplacementReport
{
    class Program
    {
        public static void Main()
        {

            //===============[Suppress warnings]======================================
#pragma warning disable CS0162
#pragma warning disable CS0164
#pragma warning disable CS0168
#pragma warning disable CS8600
#pragma warning disable CS8601
#pragma warning disable CS8602
#pragma warning disable CS8604
#pragma warning disable CA1416
            //================[Console settings]======================================
            Console.OutputEncoding = System.Text.Encoding.Unicode;

            //================[Configuration variables]==================================================================
            string strSoftwareLicenseTag = ConfigurationManager.AppSettings["SoftwareLicenseTag"];
            string strDBconnection = System.Configuration.ConfigurationManager.ConnectionStrings["DBconnectionString"].ConnectionString;

            string strProjectTitle = System.Configuration.ConfigurationManager.AppSettings["ProjectTitle"];
            string strContractTitle = ConfigurationManager.AppSettings["ContractTitle"];
            string strReportType = System.Configuration.ConfigurationManager.AppSettings["ReportType"];
            string strExcelPath = System.Configuration.ConfigurationManager.AppSettings["ExcelPath"];
            string strExcelFile = System.Configuration.ConfigurationManager.AppSettings["ExcelFile"];

            string strCheckWorksheetsExist = ConfigurationManager.AppSettings["checkWorksheetsExist"];
            string strInterpolateMissingData = ConfigurationManager.AppSettings["InterpolateMissingData"];
            string strUnits = ConfigurationManager.AppSettings["units"];

            string strReferenceWorksheet = ConfigurationManager.AppSettings["ReferenceWorksheet"];
            string strSurveyWorksheet = ConfigurationManager.AppSettings["SurveyWorksheet"];
            string strCalibrationWorksheet = ConfigurationManager.AppSettings["CalibrationWorksheet"];
            string strCurrentDisplacementWorksheet = System.Configuration.ConfigurationManager.AppSettings["CurrentDisplacementWorksheet"];
            string strHistoricDsWorksheet = System.Configuration.ConfigurationManager.AppSettings["HistoricDsWorksheet"];
            string strHistoricDhWorksheet = System.Configuration.ConfigurationManager.AppSettings["HistoricDhWorksheet"];

            string strFirstDataRow = System.Configuration.ConfigurationManager.AppSettings["FirstDataRow"];
            string strFirstDataCol = ConfigurationManager.AppSettings["FirstDataCol"];
            string strFirstOutputRow = System.Configuration.ConfigurationManager.AppSettings["FirstOutputRow"];

            string strCoordinateOrder = System.Configuration.ConfigurationManager.AppSettings["CoordinateOrder"];

            string strTimeBlockType = ConfigurationManager.AppSettings["TimeBlockType"];
            string strManualBlockStart = ConfigurationManager.AppSettings["manualBlockStart"];
            string strManualBlockEnd = ConfigurationManager.AppSettings["manualBlockEnd"];
            string strTimeOffsetHrs = ConfigurationManager.AppSettings["TimeOffsetHrs"];
            string strBlockSizeHrs = ConfigurationManager.AppSettings["BlockSizeHrs"];

            string strNoOfTimeBlocksPerReport = ConfigurationManager.AppSettings["NoOfTimeBlocksPerReport"];
            string strNoOfEpochsHistoricData = ConfigurationManager.AppSettings["NoOfEpochsHistoricData"];

            string strSendEmails = ConfigurationManager.AppSettings["SendEmails"];
            string strEmailLogin = ConfigurationManager.AppSettings["EmailLogin"];
            string strEmailPassword = ConfigurationManager.AppSettings["EmailPassword"];
            string strEmailFrom = ConfigurationManager.AppSettings["EmailFrom"];
            string strEmailRecipients = ConfigurationManager.AppSettings["EmailRecipients"];

            //================[Declare variables]===========================================================================

            // Console settings
            Console.OutputEncoding = System.Text.Encoding.Unicode;


            string strMasterWorkbookFullPath = strExcelPath + strExcelFile;

            string strTimeBlockStartLocal = "";
            string strTimeBlockEndLocal = "";
            string strTimeBlockStartUTC = "";
            string strTimeBlockEndUTC = "";

            double dblStartTimeOffset;
            double dblEndTimeOffset;

            int iRow = Convert.ToInt32(strFirstDataRow);
            int iReferenceFirstDataRow = Convert.ToInt32(strFirstDataRow);
            int iFirstOutputRow = Convert.ToInt32(strFirstOutputRow);
            int iCol = Convert.ToInt32(strFirstDataCol);
            int iNoOfEpochsHistoricData = Convert.ToInt32(strNoOfEpochsHistoricData); 
            int iNoOfTimeBlocksPerReport = Convert.ToInt32(strNoOfTimeBlocksPerReport);

            String[] strRefNo = new String[2000];
            String[] strRO1 = new String[50];
            String[] strROmeanDistances = new String[50];
            Double[,] dblNEH = new Double[2000, 3];
            String[] strName = new String[2000];
            String[] strWorksheetName = new String[6];
            string[,] strRefDistances = new String[50, 2];
            string strDateTime;

            string[,] strSensorID = new string[5000, 2];
            string[,] strPointDeltas = new string[5000, 2];

            //================[Main program]===========================================================================

            // instantiate the classes

            gnaTools gnaT = new gnaTools();
            dbAPI gnaDBAPI = new dbAPI();
            spreadsheetAPI gnaSpreadsheetAPI = new spreadsheetAPI();

            // Set the EPPlus license
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;

            // Welcome message
            gnaT.WelcomeMessage("DisplacementReport 20220913");
            gnaT.checkLicenseValidity(strSoftwareLicenseTag, strProjectTitle, strEmailLogin, strEmailPassword, strSendEmails);

            // Environment check

            Console.WriteLine("");
            Console.WriteLine("1. Check system environment");
            Console.WriteLine("");
            gnaDBAPI.testDBconnection(strDBconnection);
            Console.WriteLine("   Master workbook: " + strMasterWorkbookFullPath);
            Console.WriteLine("   Time block type: " + strTimeBlockType);

            if (strCheckWorksheetsExist == "Yes")
            {
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strReferenceWorksheet);
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strSurveyWorksheet);
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strCalibrationWorksheet);
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strCurrentDisplacementWorksheet);
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strHistoricDsWorksheet);
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strHistoricDhWorksheet);
            }
            else
            {
                Console.WriteLine("   Existance of workbook & worksheets is not checked");
            }

            //==== Prepare the time block

            switch (strTimeBlockType)
            {
                case "Manual":
                    strTimeBlockStartUTC = gnaT.convertLocalToUTC(strManualBlockStart);
                    strTimeBlockEndUTC = gnaT.convertLocalToUTC(strManualBlockEnd);
                    break;
                case "Schedule":
                    dblStartTimeOffset = -1.0 * Convert.ToDouble(strTimeOffsetHrs);
                    dblEndTimeOffset = dblStartTimeOffset - Convert.ToDouble(strBlockSizeHrs);
                    strTimeBlockStartLocal = " '" + DateTime.Now.AddHours(dblEndTimeOffset).ToString("yyyy-MM-dd HH:mm:ss") + "' ";
                    strTimeBlockEndLocal = " '" + DateTime.Now.AddHours(dblStartTimeOffset).ToString("yyyy-MM-dd HH:mm:ss") + "' ";
                    strTimeBlockStartUTC = gnaT.convertLocalToUTC(strTimeBlockStartLocal);
                    strTimeBlockEndUTC = gnaT.convertLocalToUTC(strTimeBlockEndLocal);
                    break;
                default:
                    dblStartTimeOffset = -1.0 * Convert.ToDouble(strTimeOffsetHrs);
                    dblEndTimeOffset = dblStartTimeOffset - Convert.ToDouble(strBlockSizeHrs);
                    strTimeBlockStartLocal = " '" + DateTime.Now.AddHours(dblEndTimeOffset).ToString("yyyy-MM-dd HH:mm:ss") + "' ";
                    strTimeBlockEndLocal = " '" + DateTime.Now.AddHours(dblStartTimeOffset).ToString("yyyy-MM-dd HH:mm:ss") + "' ";
                    strTimeBlockStartUTC = gnaT.convertLocalToUTC(strTimeBlockStartLocal);
                    strTimeBlockEndUTC = gnaT.convertLocalToUTC(strTimeBlockEndLocal);
                    break;
            }

            strDateTime = DateTime.Now.ToString("yyyyMMdd_HHmm");
            string strDateTimeUTC = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm");   //2022-07-26 13:45:15
            string strTimeStamp = strTimeBlockEndUTC.Replace("'", "").Trim();

            string strExportFile = strExcelPath + strContractTitle + "_" + strReportType + "_" + strDateTime + ".xlsx";
            Console.WriteLine("");

            //==== Process data ===================================================================================
            Console.WriteLine("2. Extract point names");
            string[] strPointNames = gnaSpreadsheetAPI.readPointNames(strMasterWorkbookFullPath, strSurveyWorksheet, strFirstDataRow);

            Console.WriteLine("3. Extract ProjectID");
            string strProjectID = gnaDBAPI.getProjectID(strDBconnection, strProjectTitle);

            Console.WriteLine("4. Extract LocationID");
            string[,] strNamesID = gnaDBAPI.getLocationID(strDBconnection, strProjectID, strPointNames);

            Console.WriteLine("5. Extract SensorID");
            strSensorID = gnaDBAPI.getSensorIDfromDB(strDBconnection, strNamesID);

            Console.WriteLine("6. Write SensorID to workbook");
            gnaSpreadsheetAPI.writeSensorID(strMasterWorkbookFullPath, strSurveyWorksheet, strSensorID, strFirstDataRow);

            Console.WriteLine("7. Extract mean deltas for time block");
            strPointDeltas = gnaDBAPI.getMeanDeltasFromDB(strDBconnection, strProjectTitle, strTimeBlockStartUTC, strTimeBlockEndUTC, strSensorID);

            string strBlockStart = strTimeBlockStartUTC.Replace("'", "").Trim();
            string strBlockEnd = strTimeBlockEndUTC.Replace("'", "").Trim();

            Console.WriteLine("8. Write mean deltas to master workbook");
            gnaSpreadsheetAPI.writeDeltas(strMasterWorkbookFullPath, strReferenceWorksheet, strPointDeltas, iRow, iCol, strBlockStart, strBlockEnd, strCoordinateOrder);


            Console.WriteLine("9. Calibration data");
            string strDistanceColumn = "3";
            gnaSpreadsheetAPI.populateCalibrationWorksheet(strDBconnection, strTimeBlockStartUTC, strTimeBlockEndUTC, strMasterWorkbookFullPath, strCalibrationWorksheet, strFirstOutputRow, strDistanceColumn);

            Console.WriteLine("10. Populate the historic worksheets");
            int iPrismCount = gnaSpreadsheetAPI.countPrisms(strMasterWorkbookFullPath, strReferenceWorksheet, strFirstDataRow);
            int iSourceRowStart = iFirstOutputRow;
            int iDestinationRowStart = iFirstOutputRow;
            int iHistoricDsDataColumn = gnaSpreadsheetAPI.rotateData(strMasterWorkbookFullPath, strHistoricDsWorksheet, iNoOfEpochsHistoricData, iDestinationRowStart, iPrismCount);
            int iHistoricDhDataColumn = gnaSpreadsheetAPI.rotateData(strMasterWorkbookFullPath, strHistoricDhWorksheet, iNoOfEpochsHistoricData, iDestinationRowStart, iPrismCount);
            Console.WriteLine("      Historic dS");
            gnaSpreadsheetAPI.copyHistoricData(strMasterWorkbookFullPath, strCurrentDisplacementWorksheet, strHistoricDsWorksheet, strTimeStamp, iNoOfEpochsHistoricData, iFirstOutputRow, 5, strUnits, iPrismCount, iHistoricDsDataColumn);
            Console.WriteLine("      Historic dH");
            gnaSpreadsheetAPI.copyHistoricData(strMasterWorkbookFullPath, strCurrentDisplacementWorksheet, strHistoricDhWorksheet, strTimeStamp, iNoOfEpochsHistoricData, iFirstOutputRow, 4, strUnits, iPrismCount, iHistoricDhDataColumn);

            Console.WriteLine("11. Copy workbook");
            gnaSpreadsheetAPI.copyWorkbook(strMasterWorkbookFullPath, strExportFile);

            if (strSendEmails == "Yes")
            {
                Console.WriteLine("12. Email the workbook..");

                try
                {
                    SmtpMail oMail = new SmtpMail("ES-E1582190613-00131-72B1E1BD67B73FVA-C5TC1DDC612457A3");

                    // SMTP server address
                    SmtpServer oServer = new SmtpServer("smtp.gmail.com");
                    oServer.User = strEmailLogin;
                    oServer.Password = strEmailPassword;
                    oServer.ConnectType = SmtpConnectType.ConnectTryTLS;
                    oServer.Port = 587;

                    //Set sender email address, please change it to yours
                    oMail.From = strEmailFrom;
                    oMail.To = new AddressCollection(strEmailRecipients);
                    oMail.Subject = "Displacement report: " + strProjectTitle + " (" + strDateTime + ")";
                    oMail.TextBody = "This is an automated displacement report issued by the monitoring system. Please review and forward to the client. Please do not reply to this email.";
                    oMail.AddAttachment(strExportFile);

                    SmtpClient oSmtp = new SmtpClient();
                    oSmtp.SendMail(oServer, oMail);

                }
                catch (Exception ep)
                {
                    Console.WriteLine("Failed to send email with the following error:");
                    Console.WriteLine(strEmailLogin);
                    Console.WriteLine(strEmailPassword);
                    Console.WriteLine(ep.Message);
                    //Console.ReadKey();
                }
            }

TheEnd:
            
            Console.WriteLine("\nTask complete");
            Environment.Exit(0);

        }
    }
}

