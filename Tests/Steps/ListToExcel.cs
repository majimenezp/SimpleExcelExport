using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TechTalk.SpecFlow;
using SimpleExcelExport;
using System.IO;
namespace Tests.Steps
{
    [Binding]
    public class ListToExcel
    {
        List<Person> persons;
        [Given(@"a this list of persons")]
        public void GivenAThisListOfPersons(Table table)
        {
            persons = new List<Person>();
            foreach (var row in table.Rows)
            {
                var tmp=new Person();
                tmp.Name = row["Name"];
                tmp.LastName = row["LastName"];
                tmp.BirthDay = Convert.ToDateTime(row["BirthDay"]);
                tmp.Country = row["Country"];
                tmp.Sex = (Sex)Enum.Parse(typeof(Sex), row["Sex"]);
                tmp.NumberOfChildren = Convert.ToInt32(row["NumberOfChildren"]);
                tmp.Height = Convert.ToDecimal(row["Height"]);
                persons.Add(tmp);
            }
        }
        [Then(@"export the list to a excel file located in:'(.*)'")]
        public void ThenExportTheListToAExcelFile(string fileLocation)
        {
            var result=ExportToExcel.ListToExcel<Person>(persons);
            string directory=Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            using (BinaryWriter writer = new BinaryWriter(File.Create(fileLocation)))
            {
                writer.Write(result, 0, result.Length);
                writer.Flush();
                writer.Close();
            }
        }

    }
}
