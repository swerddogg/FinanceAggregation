using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace FinanceApplication
{
    public class Doner
    {
        public string Name { get; private set; }
        public int ID { get; private set; }
        
        public Doner(string Name)
        {
            this.Name = string.IsNullOrEmpty(Name) ? "EmptyName" : Name;
            this.ID = this.GetHashCode();
        }
    }    
}
