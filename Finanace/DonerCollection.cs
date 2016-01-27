using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Xml;

namespace FinanceApplication
{    
    public class DonerCollection
    {
       // private List<Doner> doners;
       // private List<Donation> donations;

        private Dictionary<Doner, List<Donation>> data;
        private Logger logger;

        public DonerCollection(DonerCollection pastData) 
        {
        }

        public DonerCollection()
        {
            data = new Dictionary<Doner, List<Donation>>();
            logger = Logger.CreateLogger();
        }

        public void LoadDoners()
        {
            // Read the doners from the xmlInput?
            // Load them into memory
        }

        private void AddDoner(Doner newDoner)
        {
            if (IsValidDoner(newDoner) && 
                !data.ContainsKey(newDoner))
            {
                data.Add(newDoner, new List<Donation>());
            }
            else
            {
                throw new ArgumentOutOfRangeException("Doner is not acceptable. Why???");
            }
        }

        public void AddDonation(Doner doner, Donation donation)
        {
            List<Donation> singleDonation = new List<Donation>();
            singleDonation.Add(donation);
            AddDonation(doner, singleDonation);  
        }

        public void AddDonation(Doner doner, ICollection<Donation> donation)
        {
            // TODO: once we get real IDs coming from a DB, convert this to use IDs instead of Name
            var q = from d in data.Keys
                    where d.Name == doner.Name
                    select d;

            Doner newDoner;
            if (q.Count() == 0)
            {
                AddDoner(doner);
                newDoner = doner;
            }
            else
            {
                newDoner = q.First();
            }

            data[newDoner].AddRange(donation);
        }

        public List<Doner> GetAllDoners()
        {
            return data.Count > 0 ? data.Keys.ToList() : new List<Doner>();
        }

        public List<DictionaryEntry> GetAllDonersAndDonations()
        {
            List<DictionaryEntry> entries = new List<DictionaryEntry>();
            foreach (var item in data.Keys)
            {
                entries.Add(new DictionaryEntry(item, data[item]));
            }
            return entries;
        }

        public List<DictionaryEntry> GetAllDonersAndDonations(DateTime start, DateTime end)
        {
            return GetAllDonersAndDonations().Where( entry => ((Donation)entry.Value).DonationTime >= start && ((Donation)entry.Value).DonationTime < end).ToList();
        }

        public List<Donation> GetAllDonations()
        {
            List<Donation> donationList = new List<Donation>();
            foreach (var item in data.Values)
            {
                donationList.AddRange(item);                
            }
            return donationList.OrderBy(item => item.DonationTime).ToList();
        }

        public List<Donation> GetDonationsByDate(DateTime start, DateTime end)
        {
            return GetAllDonations().Where(d => d.DonationTime >= start && d.DonationTime < end).OrderBy(item => item.DonationTime).ToList();
        }

        public List<Donation> GetDonationsOfDoner(Doner doner)
        {
            return data.ContainsKey(doner) ? data[doner].OrderBy(item => item.DonationTime).ToList() : new List<Donation>();
        }

        public List<Donation> GetDonationsOfDoner(Doner doner, DateTime start, DateTime end)
        {
            return GetDonationsOfDoner(doner).Where(d => d.DonationTime >= start && d.DonationTime < end).OrderBy(item => item.DonationTime).ToList();
        }

        public List<Donation> GetDonationsOfDonerByMonth(Doner doner, int Year)
        {
            List<Donation> donations = new List<Donation>();
            for (int month = 1; month < 13; month++)
            {                  
                DateTime startDT = new DateTime(Year, month, 1);
                Donation monthDonation = new Donation(startDT); 
                List<Donation> donationsForCurrentMonth = GetDonationsOfDoner(doner, startDT, startDT.AddMonths(1));
                foreach (Donation.Category item in Enum.GetValues(typeof(Donation.Category)))
                {
                    double donationForCategory = 0.0;
                    donationForCategory =  donationsForCurrentMonth
                        .Aggregate(donationForCategory, (runningTotal, variant) => runningTotal + variant.Get(item));
                    monthDonation.Add(item, donationForCategory);
                }
                donations.Add(monthDonation);
            }

            return donations;
        }


        public List<Donation> GetDonationsByCategory(Donation.Category category)
        {
            List<Donation> donationsByCategory = new List<Donation>();
            Donation d;
            foreach (var item in GetAllDonations())
            {
                if (item.HasDonationForCategory(category))
                {
                    d = new Donation(item.DonationTime);
                    d.Add(category, item.Get(category));
                    donationsByCategory.Add(d);
                }
            }

            //var q = from dones in GetAllDonations()
            //        where dones.HasDonationForCategory(category)
            //        group dones by category into cats                    
            //        let amount = (double)cats.First().Get(category)
            //        select new Donation(cats.First().DonationTime)  into d
            //        //let nothing = d.Add(category, );
            //        where d.Add(category, amount)
            //        select d;

            return donationsByCategory.OrderBy( item => item.DonationTime ).ToList();
        }

        public List<Donation> GetDonationsByCategory(Donation.Category category, DateTime start, DateTime end)
        {
            return GetDonationsByCategory(category).Where(d => d.DonationTime >= start && d.DonationTime < end).OrderBy( d => d.DonationTime).ToList();
        }


        public void ConsolidateDoners()
        {
            string xmlFilename = System.Configuration.ConfigurationManager.AppSettings["MergeConfigFile"];
            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(xmlFilename);
            XmlNodeList merges = xDoc.SelectNodes("/DonerMerge/add");
            //logger.WriteInfo("Merging doners doncations from Config File: {0}", xmlFilename);
            foreach (XmlNode item in merges)
            {
                //logger.WriteInfo("Merging Doners --> Destination: {0}, Source: {1}", item.Attributes["destination"].Value, item.Attributes["source"].Value);
                Doner source = GetAllDoners().Where(d => d.Name.Equals(item.Attributes["source"].Value)).FirstOrDefault();
                Doner destination = GetAllDoners().Where(d => d.Name.Equals(item.Attributes["destination"].Value)).FirstOrDefault();

                if (source == null || destination == null)
                {
                    //logger.WriteError("Unable to find the destination or source in doner list");
                    continue;
                }

                if (source.Name == destination.Name)
                {
                    logger.WriteError("source and destinations match!! This must be fixed!");
                    continue;
                }

                MergeDoners(destination, source);
                logger.WriteInfo("\t completed merge..");
            }
           


        }

        public double CalculateTotal()
        {
            double total = 0.0;
            foreach (var item in data.Values)
            {
                total = (double)item.Aggregate(total, (runningTotal, variant) => runningTotal + variant.CalculateTotal()); 
            }

            return total;
        }

        public double SummarizedTotal()
        {
            double total = 0.0;
            foreach (var item in data.Values)
            {
                total = item.Aggregate(total, (runningTotal, variant) => runningTotal + variant.SummarizedTotal);
            }
            return total;
        }

        public double DepositTotal()
        {
            double total = 0.0;
            foreach (var item in data.Values)
            {
                total = item.Aggregate(total, (runningTotal, variant) => runningTotal + variant.DepositTotal);
            }
            return total;
        }

        public void MergeDonerCollection(DonerCollection mergeData)
        {
            foreach (var item in mergeData.GetAllDonersAndDonations())
            {
                this.AddDonation((Doner)item.Key, (ICollection< Donation>)item.Value);
            }
        }

        public void MergeDoners(Doner destinationDoner, Doner sourceDoner)
        {
            if (!data.ContainsKey(destinationDoner) || !data.ContainsKey(sourceDoner))
            {
                return;  // Can't merge if both aren't in the collection
            }

            AddDonation(destinationDoner, data[sourceDoner]);
            data.Remove(sourceDoner);
        }

        public void Print()
        {
            logger.WriteInfo("Number of Doners: {0}", GetAllDoners().Count);
            foreach (var d in GetAllDoners())
            {
                logger.WriteInfo("Doner: {0}, ID: {1}", d.Name, d.ID);
                foreach (var donation in GetDonationsOfDoner(d))
                {
                    donation.PrintDonations();
                }
                logger.WriteInfo("");
            }

        }

        public void PrintDoners()
        {
            logger.WriteInfo("Number of Doners: {0}", GetAllDoners().Count);
            GetAllDoners().ForEach(doner => logger.WriteInfo("Doner: {0}, ID: {1}", doner.Name, doner.ID));
        }

        public void PrintDonationCollectionDates()
        {
            IEnumerable<DateTime> collectionDates = GetAllDonations()
                .Select(d => d.DonationTime)
                .Distinct();

            foreach (var dt in collectionDates)
            {
                logger.WriteError("Collection Date: {0}", dt.ToShortDateString());
            }



        }

        private bool IsValidDoner(Doner newDoner)
        {
            return true;
        }        
    }
}
