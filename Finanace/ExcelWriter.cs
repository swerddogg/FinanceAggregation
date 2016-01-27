using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace FinanceApplication
{
    public class ExcelWriter
    {
        private string Filename;
        private Microsoft.Office.Interop.Excel.Application oApplication;
        private Microsoft.Office.Interop.Excel.Workbook oWorkbook; 
        private Microsoft.Office.Interop.Excel.Worksheet oWorksheet;
        private ReportType Report;
        private Logger logger;        
        public enum ReportType
        {
            MonthyTally,
            YearlyDonerReport
        }

        public ExcelWriter(string filename, ReportType type)
        {
            Report = type;
            oApplication = new Microsoft.Office.Interop.Excel.Application();
            oWorkbook = null;
            oWorksheet = null;
            string outputPath = Filename = System.Configuration.ConfigurationManager.AppSettings["OutputLocation"];
            
            if (type == ReportType.MonthyTally)
            {
                Filename = Path.Combine(outputPath, "MonthlyTally");
            }                
            
            if (!Directory.Exists(Filename))
                Directory.CreateDirectory(Filename);

            Filename = Path.Combine(Filename, filename);
            if (File.Exists(Filename))
                File.Delete(Filename);
            logger = Logger.CreateLogger();
        }

        private void Cleanup()
        {
            // major cleanup in order 
            if (oWorkbook != null)
            {
                oWorkbook.SaveAs(Filename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                oWorkbook.Close(false, Filename, Type.Missing);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            if (oWorksheet != null)
                Marshal.FinalReleaseComObject(oWorksheet);

            if (oWorkbook != null)
            {
                Marshal.FinalReleaseComObject(oWorkbook);
            }
            oApplication.DisplayAlerts = false;
            Marshal.FinalReleaseComObject(oApplication);
        }

        public double Write(List<Donation> donations, Doner doner)
        {
            if (Report == ReportType.MonthyTally)
                return WriteMonthlyTally(donations);
            else if (Report == ReportType.YearlyDonerReport)
                return WriteYearlyDonerReport(donations, doner);
            else
                return 0.0;
        }

        private double WriteMonthlyTally(List<Donation> donations)
        {
            double total = 0.0;
            try
            {
                if (oApplication == null)
                {
                    logger.WriteError("Unable to start an excel sheet");
                    return total;
                }
                oApplication.Visible = false;

                oWorkbook = oApplication.Workbooks.Add(Type.Missing);
                oWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)oWorkbook.Worksheets[1];
                //oWorksheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
                
                FormatFont();
                int maxRowDepth = 0;
                // Fill in all the columns (categories) with the donations 
                foreach (Donation.Category currentCategory in Enum.GetValues(typeof(Donation.Category)))
                {
                    int row = 1;
                    oWorksheet.Cells[row, (int)currentCategory] = currentCategory.ToString();
                    foreach (var item in donations)
                    {
                        string value = item.Get(currentCategory).ToString();
                        if (value.Equals("0"))
                            continue;
                        oWorksheet.Cells[++row, (int)currentCategory] = value;
                        if (currentCategory == Donation.Category.Other)
                        {
                            oWorksheet.Cells[1, (int)currentCategory + 1] = "Specify";
                            oWorksheet.Cells[row, (int)currentCategory + 1] = item.OtherCategory;
                        }
                    }
                    // keep a running total of how deep we go (rows) into the spreadsheet
                    maxRowDepth = row > maxRowDepth ? row : maxRowDepth;
                }

                int totalsRow =  CalculateTotals(2, maxRowDepth, out total);
                
                // Some formatting of row 1 (font and bold)
                oWorksheet.Rows[1, Type.Missing].Font.Size = 9;
                oWorksheet.Rows[1, Type.Missing].Font.Bold = true;

                // Totals row bold.
                oWorksheet.Cells[totalsRow, Type.Missing].Font.Bold = true;

                // Calculations at the end of the file (jeffersonville data...)
                oWorksheet.Cells[totalsRow + 3, 3] = "Total of Tithe & Offering:";
                Microsoft.Office.Interop.Excel.Range r = oWorksheet.Range[oWorksheet.Cells[totalsRow + 3, 1], oWorksheet.Cells[totalsRow + 3, 3]];
                MergeAndAlignRight(r);
                Microsoft.Office.Interop.Excel.Range TitheOffering = oWorksheet.Cells[totalsRow + 3, 4];
                TitheOffering.Formula = string.Format("=SUM({0},{1})", 
                    oWorksheet.Cells[totalsRow, (int)Donation.Category.Tithes].Address, oWorksheet.Cells[totalsRow, (int)Donation.Category.Offering].Address);
                AddBoxAroundRange(TitheOffering);
                SetCurrencyFormat(TitheOffering);

                oWorksheet.Cells[totalsRow + 1, 8] = "Jeffersonville 10%:";
                r = oWorksheet.Range[oWorksheet.Cells[totalsRow + 1, 6], oWorksheet.Cells[totalsRow + 1, 8]];
                MergeAndAlignRight(r);
                Microsoft.Office.Interop.Excel.Range Jeff10 = oWorksheet.Cells[totalsRow + 1, 10];
                Jeff10.Formula = string.Format("=({0}*0.1)", TitheOffering.Address);
                AddBoxAroundRange(Jeff10);
                SetCurrencyFormat(Jeff10);

                oWorksheet.Cells[totalsRow + 2, 8] = "Jeffersonville First Fruits:";
                r = oWorksheet.Range[oWorksheet.Cells[totalsRow + 2, 6], oWorksheet.Cells[totalsRow + 2, 8]];
                MergeAndAlignRight(r);
                Microsoft.Office.Interop.Excel.Range JeffFF = oWorksheet.Cells[totalsRow + 2, 10];
                JeffFF.Formula = string.Format("=({0}*0.5)", oWorksheet.Cells[totalsRow, (int)Donation.Category.FirstFruits].Address);
                AddBoxAroundRange(JeffFF);
                SetCurrencyFormat(JeffFF);

                oWorksheet.Cells[totalsRow + 3, 8] = "Jeffersonville Total Tithes & First Fruits:";
                r = oWorksheet.Range[oWorksheet.Cells[totalsRow + 3, 5], oWorksheet.Cells[totalsRow + 3, 8]];
                MergeAndAlignRight(r);
                oWorksheet.Cells[totalsRow + 3, 10].Formula = string.Format("=SUM({0},{1})", Jeff10.Address, JeffFF.Address);
                AddBoxAroundRange(oWorksheet.Cells[totalsRow + 3, 10]);
                SetCurrencyFormat(oWorksheet.Cells[totalsRow + 3, 10]);

                int lastRow = totalsRow + 3;
                int lastCol = (int)Donation.Category.Other + 1;
                
                // Format cells to have boxes and set print area
                AddBoxAroundEachCellInRange(oWorksheet.Range[oWorksheet.Cells[1,1], oWorksheet.Cells[totalsRow,(int)Donation.Category.Other +1]]);
                SetPrintArea(oWorksheet.Range[oWorksheet.Cells[1, 1], oWorksheet.Cells[lastRow, lastCol]]);
                //oWorksheet.Columns[Type.Missing, lastCol].PageBreak = Microsoft.Office.Interop.Excel.XlPageBreak.xlPageBreakManual;

                // set the return value equal to the total calculated
            }
            catch (Exception e )
            {
                logger.WriteError("Failure while writing excel file. Message: {0}, Stack: {1}", e.Message, e.StackTrace);
            }
            finally
            {
                Cleanup();
            }

            return total;
        }

        private void MergeAndAlignRight(Microsoft.Office.Interop.Excel.Range r)
        {
            r.Merge();
            r.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
        }

        private void AddBoxAroundRange(Microsoft.Office.Interop.Excel.Range r)
        {
            r.BorderAround(Type.Missing, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
        }

        private void AddBoxAroundEachCellInRange(Microsoft.Office.Interop.Excel.Range range)
        {
            range.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            SetCurrencyFormat(range);
        }

        private void FormatFont()
        {
            oWorksheet.Range[oWorksheet.Cells[1, 1], oWorksheet.Cells[100, 100]].Font.Name = "arial";
            oWorksheet.Range[oWorksheet.Cells[1, 1], oWorksheet.Cells[100, 100]].Font.Size = 10;
        }

        private void SetCurrencyFormat(Microsoft.Office.Interop.Excel.Range r)
        {
            string _currency = "$#,##0.00_);($#,##0.00)";
            r.NumberFormat = _currency;

        }

        private void SetPrintArea(Microsoft.Office.Interop.Excel.Range r)
        {
            oWorksheet.PageSetup.PrintArea = r.Address;
            oWorksheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
            oWorksheet.PageSetup.Zoom = false;
            oWorksheet.PageSetup.FitToPagesWide = 1;
            oWorksheet.PageSetup.FitToPagesTall = 1;
            oWorksheet.DisplayPageBreaks = true;
        }

        private double WriteYearlyDonerReportBackup(List<Donation> donations)
        {
            double total = 0.0;
            try
            {
                if (oApplication == null)
                {
                    logger.WriteError("Unable to start an excel sheet");
                    return total;
                }
                oApplication.Visible = false;

                oWorkbook = oApplication.Workbooks.Add(Type.Missing);
                oWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)oWorkbook.Worksheets[1];
                FormatFont();
                int maxRowDepth = 0;
                foreach (Donation.Category currentCategory in Enum.GetValues(typeof(Donation.Category)))
                {
                    int row = 1;
                    oWorksheet.Cells[row, (int)currentCategory] = currentCategory.ToString();
                    foreach (var item in donations)
                    {
                        oWorksheet.Cells[++row, (int)currentCategory] = item.Get(currentCategory).ToString();
                        if (currentCategory == Donation.Category.Other)
                        {
                            oWorksheet.Cells[1, (int)currentCategory + 1] = "Specify";
                            oWorksheet.Cells[row, (int)currentCategory + 1] = item.OtherCategory;
                        }
                    }
                    maxRowDepth = row > maxRowDepth ? row : maxRowDepth;
                }

                int totalsRow = CalculateTotals(2, maxRowDepth, out total);

                // Fill out the first column
                int r = 1;
                oWorksheet.Cells[r, 1] = "Date";                
                foreach (var item in donations)
                {
                    oWorksheet.Cells[++r, 1] = item.DonationTime.ToShortDateString();
                }

                // Some formatting:
                oWorksheet.Rows[1, Type.Missing].Font.Size = 9;
                oWorksheet.Rows[1, Type.Missing].Font.Bold = true;
                oWorksheet.Cells[maxRowDepth + 3, 2].Font.Bold = true;
                oWorksheet.Columns[1].Item(1).ColumnWidth = 10.71;
                total = Double.Parse(oWorksheet.Cells[maxRowDepth + 3, 2].Value);


                int lastRow = totalsRow + 3;
                int lastCol = (int)Donation.Category.Other + 1;

                // Format cells to have boxes and set print area
                AddBoxAroundEachCellInRange(oWorksheet.Range[oWorksheet.Cells[1, 1], oWorksheet.Cells[totalsRow, (int)Donation.Category.Other + 1]]);
                SetPrintArea(oWorksheet.Range[oWorksheet.Cells[1, 1], oWorksheet.Cells[lastRow, lastCol]]);
            }
            catch { }
            finally
            {
                Cleanup();
            }

            return total;
        }

        /// <summary>
        ///  Expecting to get a list of donation for a single doner
        /// </summary>
        /// <param name="donations"></param>
        /// <returns></returns>
        private double WriteYearlyDonerReport(List<Donation> donations, Doner doner)
        {
            double total = 0.0;

            try
            {
                if (oApplication == null)
                {
                    logger.WriteError("Unable to start an excel sheet");
                    return total;
                }
                oApplication.Visible = false;

                oWorkbook = oApplication.Workbooks.Add(Type.Missing);
                oWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)oWorkbook.Worksheets[1];
                FormatFont();
                int maxRowDepth = 0;
                int donationStartingRow = 10;

                MergeColumns(1, 1, 10); oWorksheet.Cells[1, 1] = "Christ Gospel Church of Tacoma WA";
                MergeColumns(2, 1, 10); oWorksheet.Cells[2, 1] = "3909 Steilacoom Blvd SW";
                MergeColumns(3, 1, 10); oWorksheet.Cells[3, 1] = "Lakewood, WA 98499";
                MergeColumns(4, 1, 10); oWorksheet.Cells[4, 1] = "(253) 584-3904";

                MergeColumns(6, 4, 8); oWorksheet.Cells[6, 4] = String.Format("As of {0:dd/MM/yyyy}", donations.Last().DonationTime.AddMonths(1).AddDays(-1.0));

                MergeColumns(8, 1, 10); oWorksheet.Cells[8, 1] = String.Format("{0}:", doner.Name);

                // Fill in all the columns (categories) with the donations 
                foreach (Donation.Category currentCategory in Enum.GetValues(typeof(Donation.Category)))
                {
                    int row = donationStartingRow;
                    oWorksheet.Cells[row, (int)currentCategory] = currentCategory.ToString();
                    foreach (var item in donations)
                    {
                        string value = item.Get(currentCategory).ToString();
                        if (value.Equals("0"))
                            value = String.Empty;
                        oWorksheet.Cells[++row, (int)currentCategory] = value;
                    }
                    // keep a running total of how deep we go (rows) into the spreadsheet
                    maxRowDepth = row > maxRowDepth ? row : maxRowDepth;
                }
                // Fill out the first column
                int rowNum = donationStartingRow;
                oWorksheet.Cells[rowNum, 1] = "Date";
                foreach (var item in donations)
                {
                    oWorksheet.Cells[++rowNum, 1] = String.Format("{0:MM/yyyy}", item.DonationTime);
                } 

                // donationStartingRow + 1 == where the actual donations start (first row is column name)
                int totalsRow = CalculateTotals(donationStartingRow+1, donationStartingRow + 1 + donations.Count , out total);

                // Some formatting of Donation Name row (font and bold)
                oWorksheet.Rows[donationStartingRow, Type.Missing].Font.Size = 9;
                oWorksheet.Rows[donationStartingRow, Type.Missing].Font.Bold = true;

                // Totals row bold.
                oWorksheet.Cells[totalsRow, Type.Missing].Font.Bold = true;

                rowNum = totalsRow + 4;
                MergeColumns(++rowNum, 1, 13); oWorksheet.Cells[rowNum, 1] = 
                    "The goods or services that Christ Gospel Church of Tacoma provided in return for your contribution consisted entirely of intangible religious benefits.";

                ++rowNum; ++rowNum;

                oWorksheet.Cells[rowNum, 1] = "Sincerely,";
                oWorksheet.Cells[++rowNum, 1] = "Treasury Department";

                int lastCol = (int)Donation.Category.Other + 1;
                SetPrintArea(oWorksheet.Range[oWorksheet.Cells[1, 1], oWorksheet.Cells[rowNum, lastCol]]);

            }
            catch (Exception e)
            {
                logger.WriteError("Failure while writing excel file. Message: {0}, Stack: {1}", e.Message, e.StackTrace);
            }
            finally
            {
                Cleanup();
            }

            return total;
        }

        private void MergeColumns(int Row, int StartingColumn, int NumberOfColumns)
        {
            Microsoft.Office.Interop.Excel.Range r = 
                oWorksheet.Range[oWorksheet.Cells[Row, StartingColumn], oWorksheet.Cells[Row, StartingColumn + NumberOfColumns]];
            r.Merge();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="startingRow"></param>
        /// <param name="endingRow"></param>
        /// <param name="TotalAmount"></param>
        /// <returns>The row which prints the total for each column</returns>
        private int CalculateTotals(int startingRow, int endingRow, out double TotalAmount)
        {
            // Create the totals row
            int colMax = Enum.GetValues(typeof(Donation.Category)).Length + 1;
            int colMin = 2;
            for (int col = colMin; col <= colMax; col++)
            {
                oWorksheet.Cells[endingRow + 2, col].Formula = 
                    string.Format("=SUM({0}:{1})", oWorksheet.Cells[startingRow, col].Address, oWorksheet.Cells[endingRow, col].Address);
            }

            // Total of all the column/rows (sums all the rows in each column, except the Date column)
            oWorksheet.Cells[endingRow + 5, (int)Donation.Category.Other - 2] = "Grand Total";
            Microsoft.Office.Interop.Excel.Range range = oWorksheet.Range[
                oWorksheet.Cells[endingRow + 5, (int)Donation.Category.Other - 3], 
                oWorksheet.Cells[endingRow + 5, (int)Donation.Category.Other - 2]];
            MergeAndAlignRight(range);
            Microsoft.Office.Interop.Excel.Range total = oWorksheet.Cells[endingRow + 5, (int)Donation.Category.Other - 1];
            total.Formula = string.Format("=SUM({0}:{1})", oWorksheet.Cells[endingRow + 2, colMin].Address, oWorksheet.Cells[endingRow + 2, colMax].Address);
            AddBoxAroundRange(total);
            SetCurrencyFormat(total);
            TotalAmount = Convert.ToDouble(total.Value);

            return endingRow + 2;
        }

    }
}
