using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace FinanceApplication
{
    public class TextWriter
    {
        private string Filename;
        private Logger logger;

        public TextWriter(string filename)
        {
            string outputPath = Filename = System.Configuration.ConfigurationManager.AppSettings["OutputLocation"];
            Filename = Path.Combine(outputPath, "YearlyDoners");
                        
            if (!Directory.Exists(Filename))
                Directory.CreateDirectory(Filename);

            Filename = Path.Combine(Filename, filename);

            if (File.Exists(Filename))
                File.Delete(Filename);
            logger = Logger.CreateLogger();
        }

        public void Write(DonerCollection donationData)
        {
            string text = "blah blah text to add";
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(String.Format("{0}: \n", donationData.GetAllDoners().FirstOrDefault().Name));
            sb.AppendLine();
            sb.AppendLine(text);
            sb.AppendLine("Your total donations for " +donationData.GetAllDonations().FirstOrDefault().DonationTime.Year + " are: ");
            sb.AppendLine(String.Format("\t{0:C}", donationData.CalculateTotal()));
            sb.AppendLine();
            sb.AppendLine();
            sb.AppendLine("Thank you for your support!");
            sb.AppendLine("--- CGC Finance Dept ---");
            using (StreamWriter file = new StreamWriter(Filename))
            {
                file.Write(sb.ToString());
            }
            
        }
    }
}
