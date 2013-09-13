using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LinqToExcel.Tests
{
    class Group
    {
        public int Id { get; set; }
        public string GroupName { get; set; }
        public List<Person> GroupMembers { get; set; }
        public List<Person> People { get; set; }
    }
}
