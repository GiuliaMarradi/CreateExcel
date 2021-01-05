using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CreateExcel
{
    class Program
    {
        static void Main(string[] args)
        {

            try
            {


                List<Package> packages =
                    new List<Package>
                        { new Package { Company = "Coho Vineyard", Weight = 25.20, TrackingNumber = 89453312L, DateOrder = DateTime.Today, HasCompleted = false },
                          new Package { Company = "Lucerne Publishing", Weight = 18, TrackingNumber = 89112755L, DateOrder = DateTime.Today, HasCompleted = false },
                          new Package { Company = "Wingtip Toys", Weight = 6.0, TrackingNumber = 299456122L, DateOrder = DateTime.Today, HasCompleted = false },
                          new Package { Company = "Adventure Works", Weight = 33.8, TrackingNumber = 4665518773L, DateOrder =  DateTime.Today.AddDays(-4), HasCompleted = true },
                          new Package { Company = "Test Works", Weight = 35.8, TrackingNumber = 4665518774L, DateOrder =  DateTime.Today.AddDays(-2), HasCompleted = true },
                          new Package { Company = "Good Works", Weight = 48.8, TrackingNumber = 4665518775L, DateOrder =  DateTime.Today.AddDays(-1), HasCompleted = true },

                        };

                List<string> headerNames = new List<string> { "Company", "Weight", "Tracking Number", "Date Order", "Completed" };

                ExcelFacade excelFacade = new ExcelFacade();
                excelFacade.Create<Package>(@"C:\temp\output1.xlsx", packages, "Packages", headerNames);

                Console.WriteLine("Completed");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            Console.Read();
        }
    }
    public class Package
    {
        public string Company { get; set; }
        public double Weight { get; set; }
        public long TrackingNumber { get; set; }
        public DateTime DateOrder { get; set; }
        public bool HasCompleted { get; set; }
    }
}
