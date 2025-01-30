using DG.XrmFramework.BusinessDomain.ServiceContext;
using DG.XrmOrg.XrmSolution.ConsoleJobs.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DG.XrmOrg.XrmSolution.ConsoleJobs.Jobs
{
    // Define a class that describes the structure of the CSV file.
    // You can use strong types.
    internal class MyCsvFile
    {
        public string Name { get; set; }
        [Header("Account ID")] //Add Header attribute if you want the CSV header to be different from the property name
        public Guid Id { get; set; }
        public Account_AccountRatingCode? AccountRatingCode { get; set; }
    }

    internal class Example : IJob
    {
        private readonly string csvFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Examples/AccountExample.csv");
        public void Run(EnvironmentConfig env)
        {
            using (var xrm = new Xrm(env.Service))
            {
                var accounts = xrm.AccountSet
                    .Where(x => x.StateCode == AccountState.Active)
                    .Select(x => new MyCsvFile
                    {
                        Name = x.Name,
                        Id = x.Id,
                        AccountRatingCode = x.AccountRatingCode,
                    })
                    .ToList();

                CsvHelper.WriteToCsv(csvFilePath, accounts); //Write to CSV, type is inferred from the input

                List<MyCsvFile> accountsFromFile = CsvHelper.ReadFromCsv<MyCsvFile>(csvFilePath); //Read from CSV

                Console.WriteLine(accountsFromFile.Count);
            }
        }
    }
}
