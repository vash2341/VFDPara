using IronPdf;
using System.Configuration;
using System.Runtime.Intrinsics.X86;
using System.Text.RegularExpressions;
using System.IO;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace VFD_Parameters
{
    internal class Program
    {
        
        /// <summary>
        ///  The main entry point for the application.  Application takes SO#, finds file on network drive, outputs all needed data for VFD Parameters
        /// </summary>
        [STAThread]
        static void Main()
        {
                   
            IronPdf.License.LicenseKey = "IRONSUITE.KPFLUEGER.CAMBRIDGEAIR.COM.28633-7706A808A9-OADFF-NLHSTIBJBTZJ-KM24AWB4OZHP-WLU3JPANAJGB-VP5DRTLYA6I7-NR4WM7DGNY4R-CEHGKA777JB7-PCO7CD-TYWKV3K3GEOKUA-DEPLOYMENT.TRIAL-V4ZGAP.TRIAL.EXPIRES.14.SEP.2023";



            //ApplicationConfiguration.Initialize();
            //Application.Run(new Form1());


            //asking user to input SO# for finding SMART paperwork
            Console.WriteLine("Please enter SO#:");
            string SO = Console.ReadLine();
            string pdfFilePath = string.Empty;
            string numericalHP = string.Empty;
            
           
            DirectoryInfo hdDirectoryInWhichToSearch = new DirectoryInfo(@"R:\Quotes\Mseries\released\");
            FileInfo[] filesInDir = hdDirectoryInWhichToSearch.GetFiles("*" + SO + "*.*");

                foreach (FileInfo foundFile in filesInDir)
                {
                    pdfFilePath = foundFile.FullName;
                    Console.Write(pdfFilePath);
                }
            

            // extracting text from pdf document
            using PdfDocument PDF = PdfDocument.FromFile(pdfFilePath);
            // get all text to put in a searchable index
            string AllText = PDF.ExtractAllText();

            

            //Search text to grab needed data
            string modelSize = AllText.Substring(AllText.LastIndexOf("MODEL:") + 7, 5);


            string blastDirection = AllText.Substring(AllText.LastIndexOf("BLAST") + 7, 9);
            if (blastDirection.Contains("UP"))
            { blastDirection = "Upblast"; }
            else if (blastDirection.Contains("Down"))
            { blastDirection = "Downblast"; }
            else
            { blastDirection = "Horizontal Blast"; }


            //get proper string to enter for gas type
            string gasType = AllText.Substring(AllText.LastIndexOf("SUPPLY:") + 8, 9);
            if (gasType.Contains("Nat"))
                {
                gasType = "Nat Gas";
                }else
                {
                gasType = "LP";
                }

            string jobName = AllText.Substring(AllText.LastIndexOf("NAME:") + 6, 15);
            string shopOrderNumber = AllText.Substring(AllText.IndexOf("Epicor") + 10, 6);
            string jobQuantity = AllText.Substring(AllText.IndexOf("QUANTITY:") + 10, 1);
            string horsePower = AllText.Substring(AllText.IndexOf("Motor -") + 7, 5);
            
            try
            {
                Regex regexObj = new Regex(@"[^\d.]");
                numericalHP = regexObj.Replace(horsePower, "");
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            string voltage = AllText.Substring(AllText.IndexOf("Control Panel - ") +16, 3);
            string phase = AllText.Substring(AllText.IndexOf("V/") + 2, 3);
            string CFM = AllText.Substring(AllText.IndexOf("MAXIMUM AIRFLOW:") + 17, 6);
            string minCFM = AllText.Substring(AllText.IndexOf("MINIMUM AIRFLOW:") + 17, 6);
            string manifoldPressureGEO = AllText.Substring(AllText.IndexOf("@ GEO:") + 7, 3);
            string manifoldPressureMAX = AllText.Substring(AllText.IndexOf("@ MAX:") + 7, 3);
            string TESP = AllText.Substring(AllText.IndexOf("TESP:") + 6, 4);


            string date = DateTime.Now.ToString(@"MM\/dd\/yyyy");
            string createdBy = "KRP";

            //create string for motor data
            string FLA = "";
            string RPM = "";


            if (numericalHP == "3" ||numericalHP == "15" ||numericalHP == "20")
            { RPM = "1765"; }
            else if(numericalHP == "2")
            { RPM = "1760"; }
            else if (numericalHP == "5")
            { RPM = "1750"; }
            else if (numericalHP == "7.5" ||numericalHP == "10" ||numericalHP == "40")
            { RPM = "1770"; }
            else if (numericalHP == "25" || numericalHP == "30" || numericalHP       == "50"  || numericalHP == "60"  || numericalHP == "75")
            { RPM = "1775"; }
            

            //460v FLA chart
            if (numericalHP  == "2"  && voltage == "460"  && AllText.Contains("ODP"))
            { FLA = "2.9"; }
            else if (numericalHP == "3" && voltage == "460" && AllText.Contains("ODP"))
            { FLA = "4.2"; }
            else if(numericalHP == "5" && voltage == "460" && AllText.Contains("ODP"))
            { FLA = "7.6"; }
            else if (numericalHP == "7.5" && voltage == "460" && AllText.Contains("ODP"))
            { FLA = "11"; }
            else if (numericalHP == "10" && voltage == "460" && AllText.Contains("ODP"))
            { FLA = "14"; }
            else if (numericalHP == "15" && voltage == "460" && AllText.Contains("ODP"))
            { FLA = "21"; }
            else if (numericalHP == "20" && voltage == "460" && AllText.Contains("ODP"))
            { FLA = "27"; }
            else if (numericalHP == "25" && voltage == "460" && AllText.Contains("ODP"))
            { FLA = "34"; }
            else if (numericalHP == "30" && voltage == "460" && AllText.Contains("ODP"))
            { FLA = "40"; }
            else if (numericalHP == "40" && voltage == "460" && AllText.Contains("ODP"))
            { FLA = "52"; }
            else if (numericalHP == "50" && voltage == "460" && AllText.Contains("ODP"))
            { FLA = "65"; }
            else if (numericalHP == "60" && voltage == "460" && AllText.Contains("ODP"))
            { FLA = "77"; }
            else if (numericalHP == "75" && voltage == "460" && AllText.Contains("ODP"))
            { FLA = "96"; }

            //208V FLA Chart
            if (numericalHP == "2" && voltage == "208" && AllText.Contains("ODP"))
            { FLA = "2.9"; }
            else if (numericalHP == "3" && voltage == "208" && AllText.Contains("ODP"))
            { FLA = "4.2"; }
            else if (numericalHP == "5" && voltage == "208" && AllText.Contains("ODP"))
            { FLA = "7.6"; }
            else if (numericalHP == "7.5" && voltage == "208" && AllText.Contains("ODP"))   
            { FLA = "11"; }
            else if (numericalHP == "10" && voltage == "208" && AllText.Contains("ODP"))
            { FLA = "14"; }
            else if (numericalHP == "15" && voltage == "208" && AllText.Contains("ODP"))
            { FLA = "21"; }
            else if (numericalHP == "20" && voltage == "208" && AllText.Contains("ODP"))
            { FLA = "27"; }
            else if (numericalHP == "25" && voltage == "208" && AllText.Contains("ODP"))
            { FLA = "34"; }
            else if (numericalHP == "30" && voltage == "208" && AllText.Contains("ODP"))
            { FLA = "40"; }
            else if (numericalHP == "40" && voltage == "208" && AllText.Contains("ODP"))
            { FLA = "52"; }
            else if (numericalHP == "50" && voltage == "208" && AllText.Contains("ODP"))
            { FLA = "65"; }
            else if (numericalHP == "60" && voltage == "208" && AllText.Contains("ODP"))
            { FLA = "77"; }
            else if (numericalHP == "75" && voltage == "208" && AllText.Contains("ODP"))
            { FLA = "96"; }

            //230V FLA Chart
            if (numericalHP == "2" && voltage == "230" && AllText.Contains("ODP"))
            { FLA = "5.8"; }
            else if (numericalHP == "3" && voltage == "230" && AllText.Contains("ODP"))
            { FLA = "8.4"; }
            else if (numericalHP == "5" && voltage == "230" && AllText.Contains("ODP"))
            { FLA = "13.2"; }
            else if (numericalHP == "7.5" && voltage == "230" && AllText.Contains("ODP"))
            { FLA = "19.6"; }
            else if (numericalHP == "10" && voltage == "230" && AllText.Contains("ODP"))
            { FLA = "25"; }
            else if (numericalHP == "15" && voltage == "230" && AllText.Contains("ODP"))
            { FLA = "36"; }
            else if (numericalHP == "20" && voltage == "230" && AllText.Contains("ODP"))
            { FLA = "48"; }
            else if (numericalHP == "25" && voltage == "230" && AllText.Contains("ODP"))
            { FLA = "60"; }
            else if (numericalHP == "30" && voltage == "230" && AllText.Contains("ODP"))
            { FLA = "72"; }
            else if (numericalHP == "40" && voltage == "230" && AllText.Contains("ODP"))
            { FLA = "98"; }
            else if (numericalHP == "50" && voltage == "230" && AllText.Contains("ODP"))
            { FLA = "114"; }
            else if (numericalHP == "60" && voltage == "230" && AllText.Contains("ODP"))
            { FLA = "136"; }
            else if (numericalHP == "75" && voltage == "230" && AllText.Contains("ODP"))
            { FLA = "170"; }

            //575V FLA Chart
            if (numericalHP == "2" && voltage == "575" && AllText.Contains("ODP"))
            { FLA = "5.8"; }
            else if (numericalHP == "3" && voltage == "575" && AllText.Contains("ODP"))
            { FLA = "8.4"; }
            else if (numericalHP == "5" && voltage == "575" && AllText.Contains("ODP"))
            { FLA = "13.2"; }
            else if (numericalHP == "7.5" && voltage == "575" && AllText.Contains("ODP"))
            { FLA = "19.6"; }
            else if (numericalHP == "10" && voltage == "575" && AllText.Contains("ODP"))
            { FLA = "25"; }
            else if (numericalHP == "15" && voltage == "575" && AllText.Contains("ODP"))
            { FLA = "36"; }
            else if (numericalHP == "20" && voltage == "575" && AllText.Contains("ODP"))
            { FLA = "48"; }
            else if (numericalHP == "25" && voltage == "575" && AllText.Contains("ODP"))
            { FLA = "60"; }
            else if (numericalHP == "30" && voltage == "575" && AllText.Contains("ODP"))
            { FLA = "72"; }
            else if (numericalHP == "40" && voltage == "575" && AllText.Contains("ODP"))
            { FLA = "98"; }
            else if (numericalHP == "50" && voltage == "575" && AllText.Contains("ODP"))
            { FLA = "114"; }
            else if (numericalHP == "60" && voltage == "575" && AllText.Contains("ODP"))
            { FLA = "136"; }
            else if (numericalHP == "75" && voltage == "575" && AllText.Contains("ODP"))
            { FLA = "170"; }

            //console write lines to check what is being pulled
            Console.WriteLine();
            Console.WriteLine("Model: " + modelSize);
            Console.WriteLine("Blast Direction: " + blastDirection);
            Console.WriteLine("Gas Type: " + gasType);
            Console.WriteLine("Job: " + jobName);
            Console.WriteLine("SO#: " + shopOrderNumber);
            Console.WriteLine("Quanitity: " + jobQuantity);
            Console.WriteLine("Date: " + date);
            Console.WriteLine("Created By: " + createdBy);
            Console.WriteLine("Horse Power: " + numericalHP);
            Console.WriteLine("Voltage: " + voltage);
            Console.WriteLine("Motor Phase: " + phase);
            Console.WriteLine("Motor RPM: " + RPM);
            Console.WriteLine("Design CFM: " + CFM);
            Console.WriteLine("Minimal CFM: " + minCFM);
            Console.WriteLine("Manifold Pressure @GEO: " + manifoldPressureGEO);
            Console.WriteLine("Manifold Pressure @MAX: " + manifoldPressureMAX);
            Console.WriteLine("Total External Static: " + TESP);
            Console.WriteLine("FLA: " + FLA);

            //open workbook for vfd paramters and write to excel workbook
            var app = new Excel.Application();
            Excel.Workbook workbook = null;

            workbook = app.Workbooks.Open(@"C:\Users\kpflueger\Desktop\Tech Stuff\VFDs\Templates\M Series VFD Worksheet Template MAIN.xlsm");
            Excel.Worksheet sheet = workbook.ActiveSheet;
            
            sheet.Range["A2"].Value = modelSize.Trim();
            sheet.Range["B2"].Value = gasType.ToString();
            sheet.Range["A3"].Value = blastDirection.ToString();
            sheet.Range["E1"].Value = jobName.ToString();
            sheet.Range["E2"].Value = shopOrderNumber.ToString();
            sheet.Range["E3"].Value = jobQuantity.ToString();
            sheet.Range["E4"].Value = date.ToString();
            sheet.Range["E5"].Value = createdBy.ToString();
            sheet.Range["B8"].Value = numericalHP.ToString();
            sheet.Range["D8"].Value = voltage.ToString();
            sheet.Range["E8"].Value = phase.ToString();
            sheet.Range["F8"].Value = FLA.ToString();
            sheet.Range["G8"].Value = RPM.ToString();
            sheet.Range["B11"].Value = CFM.ToString();
            sheet.Range["F11"].Value = TESP.ToString();
            sheet.Range["E25"].Value = manifoldPressureGEO.ToString();
            sheet.Range["E26"].Value = manifoldPressureMAX.ToString();


            double minFREQ = sheet.Range["E17"].Value;
            double geoFREQ = sheet.Range["E16"].Value;
            minFREQ = Math.Round(minFREQ, 1);
            geoFREQ = Math.Round(geoFREQ, 1);


            workbook.Save();
            workbook.Close();

            workbook = app.Workbooks.Open(@"C:\Users\kpflueger\Desktop\Tech Stuff\VFDs\Templates\ACH580 (0-10VDC).xls");
            Excel.Worksheet sheetP = workbook.ActiveSheet;

            sheetP.Range["A2"].Value = date.ToString();
            sheetP.Range["E87"].Value = "60";
            sheetP.Range["E172"].Value = FLA.ToString();
            sheetP.Range["E173"].Value = voltage.ToString();
            sheetP.Range["E175"].Value = RPM.ToString();
            sheetP.Range["E176"].Value = numericalHP.ToString();
            sheetP.Range["E86"].Value = minFREQ;
            sheetP.Range["E135"].Value = minFREQ;
            sheetP.Range["E141"].Value = geoFREQ;
            sheetP.Range["E142"].Value = geoFREQ;
            sheetP.Range["f178"].Value = manifoldPressureGEO;


            workbook.Save();
            workbook.Close();


        }



    }
}