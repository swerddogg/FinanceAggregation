using Excel;
using System;
using System.Data;
using System.IO;
using System.Linq;


namespace FinanceApplication
{
    public class ExcelFileReader : Reader  
    {
        private IExcelDataReader excelReader;
        private string Path;
        public DonerCollection Total { get; private set; }
        private Logger logger;
        private int reportYear;

        public ExcelFileReader(string Path, int ReportYear)
        {
            this.Path = Path;
            this.reportYear = ReportYear;
            Init();
            Total = new DonerCollection();
            logger = Logger.CreateLogger();
        }

        private void Init()
        {
            FileStream stream = File.Open(Path, FileMode.Open, FileAccess.Read);
            
            // Make a decision on which reader to use
            if (Path.Contains("xlsx"))
            {
                //  OpenXml Excel file (2007 format; *.xlsx)
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }
            else
            {
                //binary Excel file ('97-2003 format; *.xls)
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }

        }

        public DonerCollection Deserialize()
        {           
            excelReader.IsFirstRowAsColumnNames = true;
            DataSet result = excelReader.AsDataSet();

            DateTime donationTime = GetDonationTime();
            DonerCollection donationData = new DonerCollection();
            
            logger.WriteInfo("INFO: Reading FileName: {0}....", Path.ToString());
            foreach (DataTable sheet in result.Tables)
            {
                if (sheet.TableName.ToLowerInvariant().Contains("tally"))
                {
                    logger.WriteInfo("skipping sheet: " + sheet.TableName);
                    continue;
                }

                // Use the donation date from the worksheet name instead if it exists.
                DateTime sheetDonationTime;
                if (DateTime.TryParse(sheet.TableName, out sheetDonationTime))
                {
                    donationTime = new DateTime(reportYear, sheetDonationTime.Month, sheetDonationTime.Day);
                }

                foreach (DataRow donationRow in sheet.Rows)
                {
                    Doner currentDoner = new Doner(donationRow["Names"].ToString().Trim());
                    Donation currentDonation = new Donation(donationTime);
                    double amount = 0.0;

                    // If the row has no member name (Names) or is the date line, then skip it.
                    if (String.IsNullOrWhiteSpace(currentDoner.Name) ||
                        currentDoner.Name.ToLower().Contains("date"))
                        continue;

                    // If we have the Total Deposit line, just save this into the special variable
                    // of the Total donation (saved only in this ExcelReader) for verification
                    if (currentDoner.Name.ToLower().Contains("total deposit"))
                    {
                        if (Double.TryParse(donationRow[1].ToString(), out amount) && amount > 0.0)
                        {
                            currentDonation.DepositTotal = amount; 
                            Total.AddDonation(currentDoner, currentDonation);
                            break;
                            // This is a break because the Total Deposit row should be the last one with real 
                            // data on it. everything after is used for book-keeping...
                        }
                    }

                    foreach (DataColumn columnCategory in sheet.Columns)
                    {
                        // Trimming the column names from spreadsheet; they could contain extra spaces
                        // or they could contain "/" (hack but ok)
                        string name = columnCategory.ColumnName.ToLower().Replace(" ", String.Empty);
                        name = name.Replace("/", String.Empty);

                        string strCellValue = donationRow[columnCategory].ToString();

                        // don't bother with member column
                        if (name.Equals("names") || String.IsNullOrWhiteSpace(strCellValue))
                            continue;
                        
                        if (name.Equals("specifyother"))
                        {
                            currentDonation.OtherCategory = strCellValue;
                            continue;
                        }
                        // look for the column extracted from the spreadsheet in the known category list
                        var donationCategory = from Donation.Category c in Enum.GetValues(typeof(Donation.Category))
                                               where c.ToString().ToLower().Equals(name)
                                               select c;
                        
                        // Extract out the value from the cell
                        if (Double.TryParse(strCellValue, out amount) &&
                           (donationCategory.Count() == 1))
                        {
                            if (currentDoner.Name.ToLower().Contains("sunday school"))
                            {
                                // Always add donations for Sunday School under the Sunday School category, and keep it a running total
                                currentDonation.Add(Donation.Category.SundaySchool, currentDonation.Get(Donation.Category.SundaySchool) + amount);                                
                            }
                            else
                            {
                                currentDonation.Add(donationCategory.First(), amount);
                            }
                        }
                        else
                        {
                            logger.WriteError($"Unable to parse column or value from spreadsheet. Sheet: {sheet.TableName}, " + 
                                $"Column: {columnCategory.ColumnName}, Cell: {strCellValue}");
                                
                        }

                        // always add amount to running total
                        currentDonation.SummarizedTotal += amount;
                     
                    } 
                    
                    // Special "doner" is the total line. Save this the current object in the Total donation
                    if (currentDoner.Name.ToLower().Equals("grand totals") && currentDonation.HasDonations())
                    {
                        Total.AddDonation(currentDoner, currentDonation);
                        continue;
                    }

                    if (currentDonation.HasDonations())
                    {
                        donationData.AddDonation(currentDoner, currentDonation);
                    }
                }
            }
            logger.WriteInfo("INFO: Done reading file\n");
            return donationData;
        }
        
        private DateTime GetDonationTime()
        {
            DateTime donationTime;
            string fileName = Path.Split('\\').Last();


            if (!DateTime.TryParse(fileName.Split('.').First(), out donationTime))
            {
                // Fall back to file creation time if unable to parse date from filename.
                donationTime = File.GetCreationTime(Path);
                logger.WriteError("Unable to parse filename for date: {0}", fileName);
            }
            return donationTime;
        }
                
        public void Close()
        {
            try
            {
                excelReader.Close();
            }
            catch (Exception e)
            {
                logger.WriteWarning("Exception while closing ExcelReader: {0}", e.Message);
            }
        }

    }

}
