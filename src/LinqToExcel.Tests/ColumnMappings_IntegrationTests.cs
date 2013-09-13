using System;
using System.Linq;
using MbUnit.Framework;
using System.IO;
using log4net.Core;
using System.Data.OleDb;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Integration")]
    [TestFixture]
    public class ColumnMappings_IntegrationTests : SQLLogStatements_Helper
    {
        ExcelQueryFactory _repo;
        string _excelFileName;
        string _worksheetName;

        [TestFixtureSetUp]
        public void fs()
        {
            InstantiateLogger();
            var testDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var excelFilesDirectory = Path.Combine(testDirectory, "ExcelFiles");
            _excelFileName = Path.Combine(excelFilesDirectory, "Companies.xls");
            _worksheetName = "ColumnMappings";
        }

        [SetUp]
        public void s()
        {
            _repo = new ExcelQueryFactory();
            _repo.FileName = _excelFileName;
        }

        [Test]
        public void all_properties_have_column_mappings()
        {
            _repo.AddMapping<Company>(x => x.Name, "Company Title");
            _repo.AddMapping<Company>(x => x.CEO, "Boss");
            _repo.AddMapping<Company>(x => x.EmployeeCount, "Number of People");
            _repo.AddMapping<Company>(x => x.StartDate, "Initiation Date");

            var companies = from c in _repo.Worksheet<Company>(_worksheetName)
                            where c.Name == "Taylor University"
                            select c;

            var rival = companies.ToList().First();
            Assert.AreEqual(1, companies.ToList().Count, "Result Count");
            Assert.AreEqual("Taylor University", rival.Name, "Name");
            Assert.AreEqual("Your Mom", rival.CEO, "CEO");
            Assert.AreEqual(400, rival.EmployeeCount, "EmployeeCount");
            Assert.AreEqual(new DateTime(1988, 7, 26), rival.StartDate, "StartDate");
        }

        [Test]
        public void some_properties_have_column_mappings()
        {
            _repo.AddMapping<Company>(x => x.CEO, "Boss");
            _repo.AddMapping<Company>(x => x.StartDate, "Initiation Date");

            var companies = from c in _repo.Worksheet<Company>(_worksheetName)
                            where c.Name == "Anderson University"
                            select c;

            Company rival = companies.ToList()[0];
            Assert.AreEqual(1, companies.ToList().Count, "Result Count");
            Assert.AreEqual("Anderson University", rival.Name, "Name");
            Assert.AreEqual("Your Mom", rival.CEO, "CEO");
            Assert.AreEqual(300, rival.EmployeeCount, "EmployeeCount");
            Assert.AreEqual(new DateTime(1988, 7, 26), rival.StartDate, "StartDate");
        }

        [Test]
        public void log_warning_when_property_with_column_mapping_not_in_where_clause_when_mapped_column_doesnt_exist()
        {
            _loggedEvents.Clear();
            _repo.AddMapping<Company>(x => x.CEO, "The Big Cheese");

            var companies = from c in _repo.Worksheet<Company>(_worksheetName)
                            select c;

            companies.GetEnumerator();
            int warningsLogged = 0;
            foreach (LoggingEvent logEvent in _loggedEvents.GetEvents())
            {
                if ((logEvent.Level == Level.Warn) &&
                    (logEvent.RenderedMessage == "'The Big Cheese' column that is mapped to the 'CEO' property does not exist in the 'ColumnMappings' worksheet"))
                    warningsLogged++;
            }
            Assert.AreEqual(1, warningsLogged);
        }

        [Test]
        public void column_mappings_with_transformation()
        {
            _repo.AddMapping<Company>(x => x.IsActive, "Active", x => x == "Y");
            var companies = from c in _repo.Worksheet<Company>(_worksheetName)
                            select c;

            foreach (var company in companies)
                Assert.AreEqual(company.StartDate > new DateTime(1980, 1, 1), company.IsActive);
        }

        [Test]
        public void transformation()
        {
            //Add transformation to change the Name value to 'Looney Tunes' if it is originally 'ACME'
            _repo.AddTransformation<Company>(p => p.Name, value => (value == "ACME") ? "Looney Tunes" : value);
            var firstCompany = (from c in _repo.Worksheet<Company>(_worksheetName)
                                select c).First();

            Assert.AreEqual("Looney Tunes", firstCompany.Name);
        }

        [Test]
        public void transformation_that_returns_null()
        {
            //Add transformation to change the Name value to 'Looney Tunes' if it is originally 'ACME'
            _repo.AddTransformation<Company>(p => p.Name, value => null);
            var firstCompany = (from c in _repo.Worksheet<Company>(_worksheetName)
                                select c).First();

            Assert.AreEqual(null, firstCompany.Name);
        }

        [Test]
        public void add_second_transformation_for_new_class_same_property_name()
        {
            string testValue = "Awesome test value";
            //Add transformations to change the Name value to null for one class, testValue for another
            _repo.AddTransformation<Company>(p => p.Name, value => null);
            _repo.AddTransformation<NewCompany>(p => p.Name, value => testValue);
            var firstCompany = (from c in _repo.Worksheet<Company>(_worksheetName)
                                select c).First();
            var secondCompany = (from c in _repo.Worksheet<NewCompany>(_worksheetName)
                                select c).First();

            Assert.AreEqual(null, firstCompany.Name);
            Assert.AreEqual(testValue, secondCompany.Name);
        }

        [Test]
        public void add_second_transformation_for_new_class_same_property_name_different_property_type()
        {
            DateTime now = DateTime.Now;
            string testValue = "Awesome test value";
            //Add transformation to change the StartDate value to DateTime.Now OR testValue
            // depending on the active transform
            _repo.AddTransformation<Company>(p => p.StartDate, value => now);
            _repo.AddTransformation<NewerCompany>(p => p.StartDate, value => testValue);
            var firstCompany = (from c in _repo.Worksheet<Company>(_worksheetName)
                                select c).First();
            var secondCompany = (from c in _repo.Worksheet<NewerCompany>(_worksheetName)
                                 select c).First();

            Assert.AreEqual(now, firstCompany.StartDate);
            Assert.AreEqual(testValue, secondCompany.StartDate);
        }

        [Test]
        public void add_universal_transformation_for_all_properties_of_a_given_type()
        {
            string testValue = "Awesome test value";
            //Add transformation to change all incoming values of a certain type
            _repo.AddTransformation<string>(value => testValue);
            var firstCompany = (from c in _repo.Worksheet<Company>(_worksheetName)
                                select c).First();
            var secondCompany = (from c in _repo.Worksheet<NewCompany>(_worksheetName)
                                 select c).First();
            var thirdCompany = (from c in _repo.Worksheet<NewerCompany>(_worksheetName)
                                 select c).First();

            Assert.AreEqual(testValue, firstCompany.Name);
            Assert.AreEqual(testValue, firstCompany.CEO);
            Assert.AreEqual(testValue, secondCompany.Name);
            Assert.AreEqual(testValue, thirdCompany.Name);
            Assert.AreEqual(testValue, thirdCompany.StartDate);
        }

        [Test]
        public void add_mapping_to_link_spreadsheets_with_FK_type_references()
        {
            // Add a mapping that sets a List valued property from a query that relates another worksheet
            _repo.AddMapping<Group,Person>(g => g.GroupMembers, (g,pp) => pp.Where(p => p.GroupId == g.Id), "People");

            var groups = (from g in _repo.Worksheet<Group>("Groups") select g);
            // People table is only queried here to confirm results
            var people = (from p in _repo.Worksheet<Person>("People") select p);

            Assert.AreEqual(6, groups.Count());
            Assert.AreEqual(100, people.Count());

            var firstGroup = groups.First();
            var firstGroupMemberIds = (from p in people where p.GroupId == firstGroup.Id select p.Id).ToList();

            Assert.AreEqual(firstGroup.GroupMembers.Count, firstGroupMemberIds.Count);
            Assert.IsTrue(firstGroupMemberIds.All(id => firstGroup.GroupMembers.Any(m => m.Id == id)));
        }


        [Test]
        public void add_mapping_to_link_spreadsheets_with_FK_type_references_with_inferred_worksheet_name()
        {
            // Add a mapping that sets a List valued property from a query that relates another worksheet
            // note that no worksheet name is provided this time, as the name is inferred from the "People"
            // property of the Group class
            _repo.AddMapping<Group, Person>(g => g.People, (g, pp) => pp.Where(p => p.GroupId == g.Id));

            var groups = (from g in _repo.Worksheet<Group>("Groups") select g);
            var people = (from p in _repo.Worksheet<Person>("People") select p);

            Assert.AreEqual(6, groups.Count());
            Assert.AreEqual(100, people.Count());

            var firstGroup = groups.First();
            var firstGroupMemberIds = (from p in people where p.GroupId == firstGroup.Id select p.Id).ToList();

            Assert.AreEqual(firstGroup.People.Count, firstGroupMemberIds.Count);
            Assert.IsTrue(firstGroupMemberIds.All(id => firstGroup.People.Any(m => m.Id == id)));
        }
    }
}
