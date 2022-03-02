using ClassLibraryPopryzhenok.Utils;
using System;
using System.Linq;

namespace ClassLibraryPopryzhenok.Models
{
    public partial class Agent
    {
        public string LogoAgent
        {
            get => Logo ?? "../../Images/picture.png";
            set { }
        }

        readonly DateTime yearAgo = DateTime.Today.AddYears(-1);

        public int SalesPerYear
        {
            get
            {
                var test = AppData.db.GetSalesAfterDate(yearAgo, ID).FirstOrDefault();
                return test == null ? 0 : (int)test;
            }
        }
    }
}