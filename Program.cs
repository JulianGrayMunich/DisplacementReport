using System.Configuration;
using databaseAPI;
using EASendMail;
using GNAchartingtools;
using GNAgeneraltools;
using GNAspreadsheettools;
using GNAsurveytools;
using gnaDataClasses;
using OfficeOpenXml;
using GNA_CommercialLicenseValidator;

namespace DisplacementReport
{
    class Program
    {
        //public static string GetLicense()
        //{
        //    return license;
        //}

        //public static void Main(string license)

        static void Main()
        {

            //===============[Initial settings]======================================
#pragma warning disable CS0162
#pragma warning disable CS8600
#pragma warning disable CS8601
#pragma warning disable CS8602
#pragma warning disable CS8604
#pragma warning disable IDE0028
#pragma warning disable IDE0059



            // instantiate the classes
            gnaTools gnaT = new();
            GNAsurveycalcs gnaSurvey = new();
            dbAPI gnaDBAPI = new();
            spreadsheetAPI gnaSpreadsheetAPI = new();
            GNAchartingAPI chartingAPI = new();
            gnaDataClass gnaDC = new();


            //================[Console settings]======================================
            Console.OutputEncoding = System.Text.Encoding.Unicode;

            //================[Configuration variables]==================================================================

            string strDBconnection = System.Configuration.ConfigurationManager.ConnectionStrings["DBconnectionString"].ConnectionString;

            var config = ConfigurationManager.AppSettings;
            string strProjectTitle = config["ProjectTitle"];
            string strContractTitle = config["ContractTitle"];
            string strReportType = config["ReportType"];
            string strExcelPath = config["ExcelPath"];
            string strExcelFile = config["ExcelFile"];
            string licenseCode = config["LicenseCode"] ?? string.Empty;
            string strDebug = config["debug"];

            string strDrawCharts = config["DrawCharts"];
            string strUnits = config["units"];

            string strReferenceLineTerminalsEaNaEbNb = config["ReferenceLineTerminalsEaNaEbNb"];

            string strRootFolder = config["SystemStatusFolder"];

            string strReferenceWorksheet = config["ReferenceWorksheet"];
            string strSurveyWorksheet = config["SurveyWorksheet"];
            string strCalibrationWorksheet = config["CalibrationWorksheet"];
            string strCurrentDisplacementWorksheet = config["CurrentDisplacementWorksheet"];
            string strHistoricDsWorksheet = config["HistoricDsWorksheet"];
            string strHistoricDrWorksheet = config["HistoricDrWorksheet"];
            string strHistoricDtWorksheet = config["HistoricDtWorksheet"];
            string strHistoricDhWorksheet = config["HistoricDhWorksheet"];
            string strChartsWorksheet_dR = config["ChartsWorksheet_dR"];
            string strChartsWorksheet_dT = config["ChartsWorksheet_dT"];
            string strChartsWorksheet_dH = config["ChartsWorksheet_dH"];

            string strComputedRdT = config["computedRdT"];


            string strFirstDataRow = config["FirstDataRow"];

            string strFirstDataCol = config["FirstDataCol"];

            string strFirstOutputRow = config["FirstOutputRow"];
            string strCoordinateOrder = config["CoordinateOrder"];
            double dblDataJumpTriggerLevel = Convert.ToDouble(config["DataJumpTriggerLevel"]);

            string strCheckForOutliers = config["checkForOutliers"];

            string strTimeBlockType = config["TimeBlockType"];
            string strManualBlockStart = config["manualBlockStart"];
            string strManualBlockEnd = config["manualBlockEnd"];
            string strTimeOffsetHrs = config["TimeOffsetHrs"];
            string strBlockSizeHrs = config["BlockSizeHrs"];


            string strNoOfTimeBlocksPerReport = config["NoOfTimeBlocksPerReport"];
            string strNoOfEpochsHistoricData = config["NoOfEpochsHistoricData"];

            // Allocate recipient numbers
            var smsMobile = new string[9];
            for (int i = 0; i < smsMobile.Length; i++)
            {
                smsMobile[i] = config[$"RecipientPhone{i + 1}"] ?? string.Empty;
            }

            string strJAGstatus = config["JAGstatus"];
            string strSMStitle = config["SMSTitle"];
            string strMessage = "";

            string strSendEmails = config["SendEmails"];
            string strEmailLogin = config["EmailLogin"];
            string strEmailPassword = config["EmailPassword"];
            string strEmailFrom = config["EmailFrom"];
            string strEmailRecipients = config["EmailRecipients"];

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
            int iFirstDataRow = Convert.ToInt16(strFirstDataRow);
            int iFirstDataCol = Convert.ToInt16(strFirstDataCol);
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


            string[,] strSensorID = new string[5000, 2];
            string[,] strPointDeltas = new string[5000, 2];

            string strTab1 = "     ";
            string strTab2 = "        ";

            //================[Main program]===========================================================================

            //==== Set the EPPlus license
            //ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage.License.SetCommercial("GLKaX6q87MCqgpnTW0VeLWonJZxxBWhrLGUYIwYIap3sQwUClECEr+MXsiCn7xi5EIukcnvQCBgecfAJtn3xGgEGQzVDRjMz5wdPACsDAQEA");  //Sets your license key in code.


            //==== Validate the DSPRPT license
            LicenseValidator.ValidateLicense("DSPRPT", licenseCode);


            // Welcome message
            gnaT.WelcomeMessage("DisplacementReport 20250411");


            string strFreezeScreen = ConfigurationManager.AppSettings["freezeScreen"];

            //====  Environment check

            Console.WriteLine("");
            Console.WriteLine("1. Check system environment");
            Console.WriteLine(strTab1 + "Project: " + strProjectTitle);
            Console.WriteLine(strTab1 + "Master workbook: " + strMasterWorkbookFullPath);

            if (strFreezeScreen == "Yes")
            {
                Console.WriteLine(strTab1 + "Check database connection");
                // gnaDBAPI.testDBconnection(strDBconnection);
                Console.WriteLine(strTab1 + "Check Existance of workbook & worksheets");
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strReferenceWorksheet);
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strSurveyWorksheet);
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strCalibrationWorksheet);
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strCurrentDisplacementWorksheet);
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strHistoricDsWorksheet);

                if (strComputedRdT == "Yes")
                {
                    gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strHistoricDrWorksheet);
                    gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strHistoricDtWorksheet);
                }
                else
                {
                    Console.WriteLine(strTab1 + "Existance of historic dR & dT worksheets not checked");
                }
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strHistoricDhWorksheet);
                if (strDrawCharts == "Yes")
                {
                    gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strChartsWorksheet_dR);
                    gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strChartsWorksheet_dT);
                    gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strChartsWorksheet_dH);
                }
                else
                {
                    Console.WriteLine(strTab1 + "Existance of Charts worksheets not checked");
                }
            }
            else
            {
                Console.WriteLine(strTab1 + "Existance of workbook & worksheets is not checked");
            }


            int iNoOfPrisms = gnaSpreadsheetAPI.countPrisms(strMasterWorkbookFullPath, strCurrentDisplacementWorksheet, Convert.ToString(iFirstOutputRow), 1);


            //==== Prepare the time block
            //gnaT.checkReportingSchedule(strTimeBlockType, strProjectTitle + " DisplacementReport");

            //List<Tuple<string, string>> subBlocks = new List<Tuple<string, string>>();
            var subBlocks = new List<Tuple<string, string>>();

            switch (strTimeBlockType)
            {
                case "Historic":
                    strManualBlockStart = strManualBlockStart.Replace("'", "") + ":00";
                    strManualBlockEnd = strManualBlockEnd.Replace("'", "") + ":00";
                    strTimeBlockStartUTC = gnaT.convertLocalToUTC(strManualBlockStart).Trim();
                    strTimeBlockEndUTC = gnaT.convertLocalToUTC(strManualBlockEnd).Trim();
                    double dblBlockSizeHrs = Convert.ToDouble(strBlockSizeHrs);
                    subBlocks = gnaT.GenerateTimeBlocks(strTimeBlockStartUTC, strTimeBlockEndUTC, dblBlockSizeHrs);
                    break;
                case "Manual":
                    strManualBlockStart = strManualBlockStart.Replace("'", "") + ":00";
                    strManualBlockEnd = strManualBlockEnd.Replace("'", "") + ":00";
                    strTimeBlockStartUTC = gnaT.convertLocalToUTC(strManualBlockStart);
                    strTimeBlockEndUTC = gnaT.convertLocalToUTC(strManualBlockEnd);
                    subBlocks.Add(Tuple.Create(strTimeBlockStartUTC, strTimeBlockEndUTC));
                    break;
                case "Schedule":
                    dblStartTimeOffset = -1.0 * Convert.ToDouble(strTimeOffsetHrs);
                    dblEndTimeOffset = dblStartTimeOffset - Convert.ToDouble(strBlockSizeHrs);
                    strTimeBlockStartLocal = " '" + DateTime.Now.AddHours(dblEndTimeOffset).ToString("yyyy-MM-dd HH:mm:ss") + "' ";
                    strTimeBlockEndLocal = " '" + DateTime.Now.AddHours(dblStartTimeOffset).ToString("yyyy-MM-dd HH:mm:ss") + "' ";
                    string strLocalStartTime = DateTime.Now.AddHours(dblEndTimeOffset).ToString("yyyy-MM-dd HH:mm") + ":00";
                    string strLocalEndTime = DateTime.Now.AddHours(dblStartTimeOffset).ToString("yyyy-MM-dd HH:mm") + ":00";
                    strTimeBlockStartUTC = gnaT.convertLocalToUTC(strLocalStartTime);
                    strTimeBlockEndUTC = gnaT.convertLocalToUTC(strLocalEndTime);
                    subBlocks.Add(Tuple.Create(strTimeBlockStartUTC, strTimeBlockEndUTC));
                    break;
                default:
                    Console.WriteLine("\nError in Timeblock Type");
                    Console.WriteLine(strTab1 + "Time block type: " + strTimeBlockType);
                    Console.WriteLine(strTab1 + "Must be Manual, Schedule or Historic");
                    Console.WriteLine("\nPress key to exit..."); Console.ReadKey();
                    goto ThatsAllFolks;
                    break;
            }




            //==== Process data ===================================================================================
            Console.WriteLine("2. Extract point names");
            string[] strPointNames = gnaSpreadsheetAPI.readPointNames(strMasterWorkbookFullPath, strSurveyWorksheet, strFirstDataRow);

            Console.WriteLine("3. Extract SensorID");
            strSensorID = gnaDBAPI.getSensorIDfromDB(strDBconnection, strPointNames, strProjectTitle);

            if (strDebug == "Yes")
            {
                int Counter = 0;
                Console.WriteLine($"strProjectTitle: {strProjectTitle}");
                while (Counter < strSensorID.GetLength(0))
                {
                    string name = strSensorID[Counter, 0].Trim();
                    if (name == "NoMore") break;
                    string id = strSensorID[Counter, 1].Trim();
                    Console.WriteLine($"{Counter}  {name}  {id}");
                    Counter++;
                }

                Console.WriteLine("press key to continue...");
                Console.ReadKey();
            }

            Console.WriteLine("4. Write SensorID to workbook");
            gnaSpreadsheetAPI.writeSensorID(strMasterWorkbookFullPath, strSurveyWorksheet, strSensorID, strFirstDataRow);

            string strDateTime = DateTime.Now.ToString("yyyyMMdd_HHmm");
            string strDateTimeUTC = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss");   //2022-07-26 13:45:15
            string strTimeStamp = "";
            string strReportTime = "";
            string strExportFile = "";

            Console.WriteLine("5. Block Processing");
            Console.WriteLine(strTab1 + "Timeblock Type: " + strTimeBlockType);

            foreach (var block in subBlocks)
            {
                strTimeBlockStartUTC = block.Item1;
                strTimeBlockEndUTC = block.Item2;
                strTimeBlockStartLocal = gnaT.convertUTCToLocal(strTimeBlockStartUTC).Replace("'", "").Trim();
                strTimeBlockEndLocal = gnaT.convertUTCToLocal(strTimeBlockEndUTC).Replace("'", "").Trim();
                strReportTime = strTimeBlockEndLocal;

                strReportTime = strReportTime.Replace("-", "");
                strReportTime = strReportTime.Replace(" ", "_");
                strReportTime = strReportTime.Replace(":", "");
                strExportFile = strExcelPath + strContractTitle + "_" + strReportType + "_" + strReportTime + ".xlsx";

                strTimeStamp = strTimeBlockEndLocal + "\n(local)";
                Console.WriteLine(strTab2 + strTimeBlockStartLocal + " (local)");
                Console.WriteLine(strTab2 + strTimeBlockEndLocal + " (local)\n");


                Console.WriteLine(strTab1 + "Extract mean deltas for time block");
                strPointDeltas = gnaDBAPI.getMeanDeltasFromDB(strDBconnection, strProjectTitle, strTimeBlockStartUTC, strTimeBlockEndUTC, strSensorID);


                if (strDebug == "Yes")
                {
                    int counter = 0;
                    while (counter < strPointDeltas.GetLength(0))
                    {
                        string label = strPointDeltas[counter, 0];
                        if (label == "NoMore")
                            break;
                        Console.WriteLine($"{label}  {strPointDeltas[counter, 4]}  {strPointDeltas[counter, 5]}  counter: {strPointDeltas[counter, 6]}");
                        counter++;
                    }
                    Console.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                }

                string strBlockStart = strTimeBlockStartUTC.Replace("'", "").Trim();
                string strBlockEnd = strTimeBlockEndUTC.Replace("'", "").Trim();

                //Console.WriteLine("strBlockStart: "+ strBlockStart);
                //Console.WriteLine("strBlockEnd: " + strBlockEnd);
                //Console.WriteLine("..press key..");
                //Console.ReadKey();


                Console.WriteLine(strTab1 + "Write mean deltas to master workbook");
                gnaSpreadsheetAPI.writeDeltas(strMasterWorkbookFullPath, strReferenceWorksheet, strPointDeltas, iRow, iCol, strBlockStart, strBlockEnd, strCoordinateOrder);



                Console.WriteLine(strTab1 + "Check for gross errors");




                strMessage = gnaSpreadsheetAPI.grossAlarmCheck(strMasterWorkbookFullPath, strReferenceWorksheet, iFirstDataRow, dblDataJumpTriggerLevel);

                if ((strMessage != "OK") && (strCheckForOutliers == "Yes"))
                {
                    string strFullSMSmessage = strSMStitle + " TGR" + "\n" + strDateTime + "\n" + strMessage;
                    bool smsSuccess = gnaT.sendSMSArray(strFullSMSmessage, smsMobile);
                    Console.WriteLine(strTab1 + (smsSuccess ? "SMS sent" : "SMS failed"));

                    if (smsSuccess == true)
                    {
                        strMessage = "Displacement Report: outliers - SMS message sent";
                    }
                    else
                    {
                        strMessage = "Displacement Report: outliers - SMS message failed";
                    }
                    gnaT.updateSystemLogFile(strRootFolder, strMessage);
                }
                else if ((strMessage == "OK") && (strCheckForOutliers == "Yes"))
                {
                    Console.WriteLine(strTab1 + "No outliers");
                }
                else
                {
                    Console.WriteLine(strTab1 + "No outlier checking");
                }

                Console.WriteLine(strTab1 + "Calibration data");
                {
                    string strDistanceColumn = "3";
                    strTimeBlockStartUTC = strTimeBlockStartUTC.Replace("'", "").Trim();
                    strTimeBlockEndUTC = strTimeBlockEndUTC.Replace("'", "").Trim();
                    gnaSpreadsheetAPI.populateCalibrationWorksheet(
                        strDBconnection, strTimeBlockStartUTC, strTimeBlockEndUTC, strMasterWorkbookFullPath,
                        strCalibrationWorksheet, strFirstOutputRow, strDistanceColumn, strProjectTitle
                    );
                }

                // Here the prism data is obtained from the reference worksheet. The only time the reference worksheet is touched.
                Console.WriteLine(strTab1 + "Compute dS, dR & dT");
                List<Prism> prisms = gnaSpreadsheetAPI.computedSdRdT(strMasterWorkbookFullPath, strReferenceWorksheet, strFirstDataRow, strReferenceLineTerminalsEaNaEbNb, strComputedRdT);

                // Find first empty row
                int iFirstEmptyRow = gnaSpreadsheetAPI.findFirstEmptyRow(strMasterWorkbookFullPath, strHistoricDhWorksheet, "7", "1");
                int iLastRow = iFirstEmptyRow + 5; // to carry across any references
                int iLastCol = iNoOfEpochsHistoricData + 1;
                // Find first empty col
                int iFirstEmptyCol = gnaSpreadsheetAPI.findFirstEmptyColumn(strMasterWorkbookFullPath, strHistoricDhWorksheet, "5", "3");
                int iStartRow = 5;
                int iStartCol = 4;
                // Bump columns left if necessary
                Console.WriteLine(strTab1 + "Delete Historic columns: " + iNoOfEpochsHistoricData + " columns");
                if (iFirstEmptyCol > iLastCol)
                {
                    Console.WriteLine(strTab2 + strHistoricDsWorksheet);
                    gnaSpreadsheetAPI.ShiftCellsLeft(strMasterWorkbookFullPath, strHistoricDsWorksheet, iStartRow, iLastRow, iStartCol, iFirstEmptyCol);
                    Console.WriteLine(strTab2 + strHistoricDhWorksheet);
                    gnaSpreadsheetAPI.ShiftCellsLeft(strMasterWorkbookFullPath, strHistoricDhWorksheet, iStartRow, iLastRow, iStartCol, iFirstEmptyCol);
                    if (strComputedRdT == "Yes")
                    {
                        Console.WriteLine(strTab2 + strHistoricDrWorksheet);
                        gnaSpreadsheetAPI.ShiftCellsLeft(strMasterWorkbookFullPath, strHistoricDrWorksheet, iStartRow, iLastRow, iStartCol, iFirstEmptyCol);
                        Console.WriteLine(strTab2 + strHistoricDtWorksheet);
                        gnaSpreadsheetAPI.ShiftCellsLeft(strMasterWorkbookFullPath, strHistoricDtWorksheet, iStartRow, iLastRow, iStartCol, iFirstEmptyCol);
                    }
                }
                else
                {
                    Console.WriteLine(strTab2 + "No columns to be deleted");
                }


                Console.WriteLine(strTab1 + "Write to " + strCurrentDisplacementWorksheet);
                //Console.WriteLine(strTab1+"Write to worksheets");
                gnaSpreadsheetAPI.writedSdRdTtoWorksheets(strMasterWorkbookFullPath, prisms, strCurrentDisplacementWorksheet, strHistoricDrWorksheet, strHistoricDtWorksheet, strHistoricDhWorksheet, strHistoricDsWorksheet, strComputedRdT, iFirstDataRow, iFirstOutputRow, strTimeBlockEndLocal);

                // Get timestamp
                strTimeStamp = strTimeBlockEndLocal + "\n(local)";
                int iTargetCol = gnaSpreadsheetAPI.findFirstEmptyColumn(strMasterWorkbookFullPath, strHistoricDsWorksheet, "5", "2");

                // Copy values to historic worksheets
                Console.WriteLine(strTab1 + "Copy data to historic worksheets");
                int iEndRow = iFirstOutputRow + iNoOfPrisms - 1;
                iStartRow = iFirstOutputRow;
                Console.WriteLine(strTab2 + strHistoricDsWorksheet);
                gnaSpreadsheetAPI.copyColumnRange(strMasterWorkbookFullPath, strCurrentDisplacementWorksheet, 4, strHistoricDsWorksheet, iTargetCol, iStartRow, iEndRow, strTimeStamp);
                Console.WriteLine(strTab2 + strHistoricDhWorksheet);
                gnaSpreadsheetAPI.copyColumnRange(strMasterWorkbookFullPath, strCurrentDisplacementWorksheet, 7, strHistoricDhWorksheet, iTargetCol, iStartRow, iEndRow, strTimeStamp);

                if (strComputedRdT == "Yes")
                {
                    Console.WriteLine(strTab2 + strHistoricDrWorksheet);
                    gnaSpreadsheetAPI.copyColumnRange(strMasterWorkbookFullPath, strCurrentDisplacementWorksheet, 5, strHistoricDrWorksheet, iTargetCol, iStartRow, iEndRow, strTimeStamp);
                    Console.WriteLine(strTab2 + strHistoricDtWorksheet);
                    gnaSpreadsheetAPI.copyColumnRange(strMasterWorkbookFullPath, strCurrentDisplacementWorksheet, 6, strHistoricDtWorksheet, iTargetCol, iStartRow, iEndRow, strTimeStamp);
                }

            }

            //Console.WriteLine("6. Compute displacement factor");
            //Console.WriteLine(strTab1 + "No");
            //gnaSpreadsheetAPI.computePositiveDisplacementFactor(strMasterWorkbookFullPath, strReferenceWorksheet, strPositiveDisplacementBearing, strFirstDataRow);

            //Console.WriteLine("7. Copy workbook");
            gnaSpreadsheetAPI.copyWorkbook(strMasterWorkbookFullPath, strExportFile);

            //Console.WriteLine("8. Reset the master workbook");
            iFirstOutputRow = Convert.ToInt16(strFirstOutputRow);
            int iLastOutputRow = iFirstOutputRow + iNoOfPrisms - 1;
            gnaSpreadsheetAPI.resetMasterWorkbook(strMasterWorkbookFullPath, strCurrentDisplacementWorksheet, iFirstOutputRow, 4, iLastOutputRow, 7);

            Console.WriteLine("9. Draw charts");
            strDateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

            if (strDrawCharts == "Yes")
            {

                var chart = new ChartSettings();

                chart.chartWorksheet = "Charts";
                chart.YseriesRow = 7;   // vertical
                chart.XseriesRow = 6;   // horizontal
                chart.XAxisTitle = "Date";
                chart.YAxisTitle = "Displacement (m)";
                chart.firstDataCol = 3;
                chart.lastDataCol = 16;
                chart.firstDataRow = 7;
                chart.lastDataRow = 17;
                chart.legendCol = "A";


                Console.WriteLine(strTab1 + "Historic displacement");
                chart.chartX = 250;
                chart.chartY = 300;
                chart.dataWorksheet = "Historic Displacement";
                chart.chartName = "Displacement";
                chart.chartTitle = "Horizontal Displacement \nPast 14 days";
                chart.YAxisTitle = "Displacement (m)";
                chart.XAxisTitle = "Days";
                chart.YAxisMaxValue = .01;
                chart.YAxisMinValue = -0.01;
                chartingAPI.drawMultiSeriesChart(strExportFile, chart);


                Console.WriteLine(strTab1 + "Horizontal settlement");
                chart.chartX = 1100;
                chart.chartY = 300;
                chart.dataWorksheet = "Historic Settlement";
                chart.chartName = "Settlement";
                chart.chartTitle = "Vertical Settlement \nPast 14 days";
                chart.YAxisMaxValue = .01;
                chart.YAxisMinValue = -0.01;
                chartingAPI.drawMultiSeriesChart(strExportFile, chart);

            }
            else
            {
                Console.WriteLine(strTab1 + "No charts");
            }

            Console.WriteLine("10. Email the workbook");
            if (strSendEmails == "Yes")
            {
                try
                {
                    strMessage = "This is an automated displacement report issued by the monitoring system. Please review and forward to the client. Please do not reply to this email.";
                    strMessage = gnaT.addCopyright("DsplRpt", strMessage);
                    string license = gnaT.commercialSoftwareLicense("email");

                    SmtpMail oMail = new(license)
                    {
                        From = strEmailFrom,
                        To = new AddressCollection(strEmailRecipients),
                        Subject = "Displacement report: " + strProjectTitle + " (" + strDateTime + ")",
                        TextBody = strMessage
                    };

                    // SMTP server address
                    SmtpServer oServer = new("smtp.gmail.com")
                    {
                        User = strEmailLogin,
                        Password = strEmailPassword,
                        ConnectType = SmtpConnectType.ConnectTryTLS,
                        Port = 587
                    };


                    oMail.AddAttachment(strExportFile);
                    SmtpClient oSmtp = new();
                    oSmtp.SendMail(oServer, oMail);

                    strMessage = "Displacement report: " + strProjectTitle + " (emailed)";

                    gnaT.updateSystemLogFile(strRootFolder, strMessage);
                    gnaT.updateReportTime("DSPRPT");
                    Console.WriteLine(strTab1 + "email sent & logs updated");
                }
                catch (Exception ep)
                {
                    Console.WriteLine(strTab1 + "\nFailed to send email with the following error:");
                    Console.WriteLine(strEmailLogin);
                    Console.WriteLine(strEmailPassword);
                    Console.WriteLine(ep.Message);
                    //Console.ReadKey();
                }
            }
            else
            {
                Console.WriteLine(strTab1 + "No email sent");
            }

ThatsAllFolks:
            gnaT.freezeScreen(strFreezeScreen);
            Console.WriteLine("\nTask complete");
            Environment.Exit(0);

        }
    }
}

