using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Configuration;
//using Excel;
//using OfficeOpenXml; 
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Globalization;

namespace FinanceApplication
{
    class Program
    {
        static Logger logger;
        static int reportingYear = int.Parse(ConfigurationManager.AppSettings["ReportingYear"]);
        static List<string> reportTypes = ConfigurationManager.AppSettings["ReportType"].Split(';').ToList();
        static void Main(string[] args)
        {
            DonerCollection data;
            try
            {
                logger = Logger.CreateLogger(true);
                //Test();
                // string filePath = @"TestData\";
                string filePath = System.Configuration.ConfigurationManager.AppSettings["SpreadSheetLocation"];
                string[] files = Directory.GetFiles(filePath, "*.xls*");
                data = new DonerCollection();
                ReadExcel(files, data);
                //TestMerge();
                //TestLogger();
                //TestFileNames(files);
                //TestDonationQueries();
                data.PrintDoners();
               // TestPrintDonationsOfAaron(data);
                data.ConsolidateDoners();
                data.PrintDoners();
                data.PrintDonationCollectionDates();
                if (reportTypes.Contains("Donors"))
                    CreateReportForDoners(data);

                if (reportTypes.Contains("Monthly"))
                    CreateMonthlyReports(data);
                //TestExcelWriter(ref data);
               
            }
            catch (Exception e)
            {
                logger.WriteError("Exception occurred during execution: \n {0}, {1}", e.Message, e.StackTrace);
            }
            finally
            {
                logger.Close();
            }
        }

        private static void CreateReportForDoners(DonerCollection data)
        {
            List<Doner> doners = data.GetAllDoners();
            foreach (var item in doners)
            {
                logger.WriteInfo("Processing Doner: {0} ....", item.Name);
                ExcelWriter writer = new ExcelWriter(String.Format("{0}_Donations{1}.xlsx", item.Name, reportingYear), 
                    ExcelWriter.ReportType.YearlyDonerReport);
                DonerCollection donationsByDoner = new DonerCollection();
                donationsByDoner.AddDonation(item, data.GetDonationsOfDonerByMonth(item, reportingYear));
                writer.Write(donationsByDoner.GetAllDonations(), item);
            }
        }


        private static void CreateMonthlyReports(DonerCollection data)
        {
            List<Doner> doners = data.GetAllDoners();
            for (DateTime start = new DateTime(reportingYear, 1, 1); start.Year < reportingYear+1; start = start.AddMonths(1))
            //for (DateTime start = new DateTime(2010, 1, 1); start.Month < 2; start = start.AddMonths(1))
            {
                DateTime end = start.AddMonths(1);
                logger.WriteInfo("Processing Month: {0} ....", CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName( start.Month ));
                var monthData = data.GetDonationsByDate(start, end);
                if (!monthData.Any())
                {
                    logger.WriteWarning("No data found for this month");
                    continue;
                }
                ExcelWriter writer = new ExcelWriter(String.Format("{0}_{1}_TallyReport.xlsx", 
                    CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName( start.Month ), start.Year.ToString()),
                    FinanceApplication.ExcelWriter.ReportType.MonthyTally);
                logger.WriteInfo("Done. Total = {0:C}", writer.Write(monthData, new Doner("nobody")));
                // TEMP CODE FOR DEBUGGING
                //data.GetDonationsByDate(start, end).ForEach(d => d.PrintDonations());

            }
        }

        private static void TestPrintDonationsOfAaron(DonerCollection data)
        {
            List<Doner> doner = data.GetAllDoners().Where(d => d.Name.Contains("Aaron")).ToList();
            foreach (var item in doner)
            {
                logger.WriteInfo("Printing Donations for {0}", item.Name);
                data.GetDonationsOfDoner(item).ForEach(donation => donation.PrintDonations());
                double donationTotal = 0.0;
                logger.WriteInfo("Total of all donations is ${0}", data.GetDonationsOfDoner(item).Aggregate(donationTotal, (total, variant) => total + variant.CalculateTotal()));
                logger.WriteInfo("Processing Doner: {0} ....", item.Name);
                ExcelWriter writer = new ExcelWriter(String.Format("{0}_Donations2010.xlsx", item.Name), FinanceApplication.ExcelWriter.ReportType.YearlyDonerReport);
                writer.Write(data.GetDonationsOfDoner(item), item);    
            }
        }

        private static void TestPrintDonationsOfAaronEx(DonerCollection data)
        {
            List<Doner> doner = data.GetAllDoners().Where(d => d.Name.Contains("Aaron")).ToList();
            foreach (var item in doner)
            {
                logger.WriteInfo("Printing Donations for {0}", item.Name);
                data.GetDonationsOfDoner(item).ForEach(donation => donation.PrintDonations());
                double donationTotal = 0.0;
                logger.WriteInfo("Total of all donations is ${0}", data.GetDonationsOfDoner(item).Aggregate(donationTotal, (total, variant) => total + variant.CalculateTotal()));
                logger.WriteInfo("Processing Doner: {0} ....", item.Name);
                ExcelWriter writer = new ExcelWriter(String.Format("{0}_Donations2010.xlsx", item.Name), FinanceApplication.ExcelWriter.ReportType.YearlyDonerReport);
                writer.Write(data.GetDonationsOfDoner(item), item);

                logger.WriteInfo("Printing Donations by Month......");
                data.GetDonationsOfDonerByMonth(item, 2010).ForEach( d => d.PrintDonations() );
            }
        }
        private static void ReadExcel(string[] files, DonerCollection data)
        {

            List<string> ErrorFiles = new List<string>();
            foreach (var item in files)
            {
                logger.WriteInfo("\n___________________________________START READ__________________________________________");
                ExcelFileReader xRead = new ExcelFileReader(item, reportingYear);
                DonerCollection currentSheet = xRead.Deserialize();
                //currentSheet.Print();

                double currentFileCalcTotal = xRead.Total.CalculateTotal();
                double currentFileSumTotal = xRead.Total.SummarizedTotal();
                double currentFileDepTotal = xRead.Total.DepositTotal();
                double collectionCalcTotal = currentSheet.CalculateTotal();
                double collectionSumTotal = currentSheet.SummarizedTotal();
                
                // Run some validation to compare all the donations received per user against the "Totals" row and also
                // against a running total of values read (regardless of their category)
                if (!CompareDbls(currentFileCalcTotal, collectionCalcTotal) ||
                    !CompareDbls(collectionCalcTotal, currentFileDepTotal) ||
                    !CompareDbls(currentFileDepTotal, currentFileSumTotal) ||
                    !CompareDbls(currentFileSumTotal, collectionSumTotal))
                {
                    logger.WriteInfo("\tSpreadsheet       Calculated Total:  {0:C}", currentFileCalcTotal);
                    logger.WriteInfo("\tCollections Data  Calculated Total:  {0:C}", collectionCalcTotal);
                    logger.WriteInfo("\tSpreadsheet       Total Deposit Row: {0:C}", currentFileDepTotal);
                    logger.WriteInfo("\tSpreadsheet       Grand Totals Row:  {0:C}", currentFileSumTotal);
                    logger.WriteInfo("\tCollections Data  Grant Totals Row:  {0:C}", collectionSumTotal);
                    logger.WriteError("Totals do not add up correctly. Check the file for consistencies: {0}", item);
                    ErrorFiles.Add(item);
                }
                logger.WriteInfo("\n");
                xRead.Close();

                double dataCollectionCalcTotalBefore = data.CalculateTotal();
                data.MergeDonerCollection(currentSheet);

                double dataCollectionCalcTotalAfter = data.CalculateTotal();


                // One last validation to make sure the data merged into the overall collection maintains a consistent summary
                if (!CompareDbls(dataCollectionCalcTotalAfter, (dataCollectionCalcTotalBefore + collectionCalcTotal)))
                {
                    logger.WriteInfo("INFO: Post Merge");
                    logger.WriteInfo("\tCalculated Total in Data Collection before Merge: {0}", dataCollectionCalcTotalBefore);
                    logger.WriteInfo("\tCalculated Total from the new current Collection: {0}", collectionCalcTotal);
                    logger.WriteInfo("\tCalculated Total in Data Collection after  Merge: {0}", dataCollectionCalcTotalAfter);
                    logger.WriteInfo("Problem occurred after merge. Data did not propagate correctly");
                }

                logger.WriteInfo("___________________________________FINISH READ___________________________________________\n");


            }
            logger.WriteInfo("___________________________________START SUMMARY OF ALL DATA______________________________\n");
            //data.Print();
            logger.WriteInfo("___________________________________END SUMMARY____________________________________________\n");


            logger.WriteInfo("Summary of Results");
            logger.WriteInfo("Below is the list of filenames which had a problem importing: ");
            ErrorFiles.ForEach(file => logger.WriteInfo("\t {0}", file));
            if (ErrorFiles.Count() == 0)
                logger.WriteInfo("\t No Files...");
        }

        private static bool CompareDbls(double first, double second)
        {
            return Math.Abs(first - second) < 0.001;
        }
        
        private static void Test()
        {

            Doner me = new Doner("Aaron Swerdlin");

            Donation money = new Donation();
            money.Add(Donation.Category.Tithes, 22.00);
            money.Add(Donation.Category.FirstFruits, 23.900);
            money.Add(Donation.Category.Alms, 20.10);

            money.PrintDonations();

            money.Add(Donation.Category.Tithes, 1100.00);
            money.PrintDonations();
            logger.WriteInfo("Totals: {0}: ", money.CalculateTotal());
        }

        private static void TestMerge()
        {
            string[] files = { @"TestData\1-2-2009.xls" };
            DonerCollection data = new DonerCollection();
            ReadExcel(files, data);

            logger.WriteInfo("*************TIME TO MERGE******************");
            List<Doner> doners = data.GetAllDoners();
            data.MergeDoners(doners.First(), doners.Last());

            data.Print();



        }

        private static void TestFileNames(string[] files)
        {           
            foreach (var item in files)
            {
                logger.WriteInfo("Filename: {0}, Date: {1}", item, GetDonationTime(item).ToShortDateString());
            }
        }

        private static DateTime GetDonationTime(string Path)
        {
            DateTime donationTime;
            string fileName = Path.Split('\\').Last();


            if (!DateTime.TryParse(fileName.Split('.').First(), out donationTime))
            {
                // Fall back to file creation time if unable to parse date from filename.
                donationTime = File.GetCreationTime(Path);
            }
            return donationTime;
        }

        private static void TestLogger()
        {
            logger.WriteInfo("Simple input with no parameters");
            logger.WriteInfo("One parameter {0} in the middle", "HERE");
            logger.WriteInfo("Two parameters named: {0} and {1}", "A", "B");

            logger.WriteError("Simple input with no parameters");
            logger.WriteWarning("One parameter {0} in the middle", "HERE");
            logger.WriteError("Two parameters named: {0} and {1}", "A", "B");

        }

        private static void TestDonationQueries()
        {
            string[] files = { @"C:\temp\TestData\03-07-2010.xls",
                             @"C:\temp\TestData\03-8-2010.xls",
                             @"C:\temp\TestData\2-28-2010.xls",
                             @"C:\temp\TestData\8-1-2010.xls",
                            @"C:\temp\TestData\09-01-2010.xls"};
            DonerCollection data = new DonerCollection();
            ReadExcel(files, data);

            logger.WriteInfo("*************ALL DATA START******************");
            data.Print();
            logger.WriteInfo("*************ALL DATA END ******************");

            logger.WriteInfo("*************ALL DONERS START******************");
            data.PrintDoners();
            logger.WriteInfo("*************ALL DONERS END ******************");


            logger.WriteInfo("*************ALL DONATIONS FOR FIRST START******************");
            data.GetDonationsOfDoner(data.GetAllDoners().First()).ForEach( d => d.PrintDonations() );
            logger.WriteInfo("*************ALL DONATIONS FOR FIRST END ******************");

            //logger.WriteInfo("*************ALL DONATIONS FOR FIRST START******************");
            //data.GetCategoryDonations(Donation.Category.Tithes).ForEach(d => d.PrintDonations());
            //logger.WriteInfo("*************ALL DONATIONS FOR FIRST END ******************");

            logger.WriteInfo("*************ALL TITHES START******************");
            data.GetDonationsByCategory(Donation.Category.Tithes).ForEach(d => d.PrintDonations());
            logger.WriteInfo("*************ALL TITHES END ******************");


            logger.WriteInfo("*************RANGE TITHES START******************");
            data.GetDonationsByCategory(Donation.Category.Tithes, new DateTime(2010, 02, 01), new DateTime(2010, 04, 01)).ForEach(d => d.PrintDonations());
            logger.WriteInfo("*************RANGE TITHES END ******************");


            logger.WriteInfo("*************ALL DONATIONS FOR FEB START******************");
            data.GetAllDonersAndDonations(new DateTime(2010, 02, 01), new DateTime(2010, 02, 29))
                .ForEach(d => 
                {
                    Console.WriteLine( "Name: {0}", ((Doner)d.Key).Name);
                    ((Donation)d.Value).PrintDonations(); 
                } );
            logger.WriteInfo("*************ALL DONATIONS FOR FEB END ******************");
        }

        private static bool TestExcelWriter(ref DonerCollection data)
        {
            List<Doner> doners = data.GetAllDoners().Where( d => d.Name.ToLower().Contains("aaron")).ToList();
            foreach (var item in doners)
            {
                logger.WriteInfo("Processing Doner: {0} ....", item.Name);
                ExcelWriter writer = new ExcelWriter(String.Format("{0}_Donations{1}.xlsx", item.Name, reportingYear), 
                    FinanceApplication.ExcelWriter.ReportType.YearlyDonerReport);
                writer.Write(data.GetDonationsOfDoner(item), item);
            }
            return true;
        }

        
    }
}
