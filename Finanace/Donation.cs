using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinanceApplication
{
    public class Donation
    {
        //public enum Category
        //{
        //    Tithes,
        //    FirstFruits,
        //    Missions,
        //    Alms,
        //    BuildingFund,
        //    Benevolence,
        //    PastorFund,
        //    Offering,
        //    Youth,
        //    Other
        //}

        public enum Category
        {
            Tithes = 2,
            FirstFruits,
            Missions,
            Alms,
            Gifts,
            Building,
            RoyalRegiments,
            Elders,
            Literature,
            PastorTravelEd,
            Benevolence,
            Offering,
            SundaySchool,
            Youth,
            Other
        }

        public string OtherCategory { get; set; } // aka "Please Specify"

        // slightly hacky, need a way to keep a running total as entries are added when reading from 
        // excel. This is to ensure the columns in excel match up with what we're expecting in the code.
        // if there's a difference in this Total vs Calculate Total then we know there was an error
        // in reading in from excel...
        public double SummarizedTotal { get; set; }

        // Taken from the row titled "Deposit Title" in the excel sheet.
        public double DepositTotal { get; set; }
        public DateTime DonationTime { get; private set; } 

        private Dictionary<Category, double> donations;
        private Logger logger;
        public Donation(DateTime DonationTime)
        {
            this.DonationTime = DonationTime;
            donations = new Dictionary<Category, double>();
            logger = Logger.CreateLogger(); 
        }

        public Donation() : this (DateTime.Now)
        {
        }

        /// <returns>true if category already existed</returns>
        public bool Add(Category category, double amount)
        {
            bool ret = false;

            if (HasDonationForCategory(category))
            {
                ret = true;
                donations.Remove(category);
            }

            donations.Add(category, amount);
            return ret;
        }

        public double Get(Category category)
        {
            return HasDonationForCategory(category) ? donations[category] : 0.0;
        }

        public bool HasDonationForCategory(Category category)
        {
            return donations.ContainsKey(category);
        }

        public bool HasDonations()
        {
            return donations.Count() > 0;
        }

        public void PrintDonations()
        {
            if (donations.Count() > 0)
                logger.WriteInfo("Donations ({0}): ", this.DonationTime.ToLongDateString());
            foreach ( var item in donations )
            {
                logger.WriteInfo("\t {0}: ${1:0.00}", item.Key.ToString(), item.Value);
                if (item.Key == Category.Other)
                {
                    logger.WriteInfo("\t\tOther Category: {0}", OtherCategory);
                }
            }
        }

        public double CalculateTotal()
        {
            return donations.Count > 0 ? 
                donations.Values.Aggregate((double currentTotal, double donation) => currentTotal + donation) :
                0.0; 
        }
        
    }
}
