using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    public class GameAndApp
    { 
        public string name { get; set; }
        public string urlApp { get; set; }
        public string ageApp { get; set; }
        public string pricing { get; set; }
        public string size { get; set; }
        public string category { get; set; }
        public string description { get; set; }
        public string developer { get; set; }

        public GameAndApp(string name, string urlApp, string ageApp, string pricing, string size, string category, string description, string developer)
        {
            this.name = name;
            this.urlApp = urlApp;
            this.ageApp = ageApp;
            this.pricing = pricing;
            this.size = size;
            this.category = category;
            this.description = description;
            this.developer = developer;
        }
        public GameAndApp()
        {
        }
    }
}
