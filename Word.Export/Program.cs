using System;
using System.Collections.Generic;
using System.IO;

namespace Word.Export
{
    class Program
    {
        static void Main()
        {
            var enters = new List<Enter>
            {
                new Enter{Order = 1, Name = "杨旭升", EnterDate = DateTime.Parse("2015-01-26"), IdNum = "123456789123456789", Job = "木工", IsRecord = true, HasContract = false},
                new Enter{Order = 2, Name = "沈建昌", EnterDate = DateTime.Parse("2015-01-26"), IdNum = "123456789123456789", Job = "木工", IsRecord = false, HasContract = true},
                new Enter{Order = 3, Name = "胡永成", EnterDate = DateTime.Parse("2015-01-26"), IdNum = "123456789123456789", Job = "木工", IsRecord = true, HasContract = false},
                new Enter{Order = 4, Name = "葛  军", EnterDate = DateTime.Parse("2015-01-27"), IdNum = "123456789123456789", Job = "木工", IsRecord = false, HasContract = true},
                new Enter{Order = 5, Name = "吴亚亦", EnterDate = DateTime.Parse("2015-01-27"), IdNum = "123456789123456789", Job = "木工", IsRecord = true, HasContract = false},
                new Enter{Order = 6, Name = "王国祥", EnterDate = DateTime.Parse("2015-01-27"), IdNum = "123456789123456789", Job = "木工", IsRecord = false, HasContract = true},
                new Enter{Order = 7, Name = "马群东", EnterDate = DateTime.Parse("2015-01-28"), IdNum = "123456789123456789", Job = "木工", IsRecord = true, HasContract = false},
                new Enter{Order = 8, Name = "何  叶", EnterDate = DateTime.Parse("2015-01-28"), IdNum = "123456789123456789", Job = "木工", IsRecord = false, HasContract = true},
                new Enter{Order = 9, Name = "王  强", EnterDate = DateTime.Parse("2015-01-28"), IdNum = "123456789123456789", Job = "木工", IsRecord = true, HasContract = false},
                new Enter{Order = 10, Name = "范双元", EnterDate = DateTime.Parse("2015-01-29"), IdNum = "123456789123456789", Job = "木工", IsRecord = false, HasContract = true}
            };

            var departures = new List<Departure>
            {
                new Departure{Order = 1,  Name = "杨旭升", DepartureDate =  DateTime.Parse("2015-01-26"),  IdNum = "123456789123456789", Job = "木工", WorkDays = 10, IsSettled = true, IsCommited = false},
                new Departure{Order = 2,  Name = "沈建昌", DepartureDate =  DateTime.Parse("2015-01-27"),  IdNum = "123456789123456789", Job = "木工", WorkDays = 10, IsSettled = false, IsCommited = true},
                new Departure{Order = 3,  Name = "胡永成", DepartureDate =  DateTime.Parse("2015-01-28"),  IdNum = "123456789123456789", Job = "木工", WorkDays = 10, IsSettled = true, IsCommited = false},
                new Departure{Order = 4,  Name = "葛  军", DepartureDate =  DateTime.Parse("2015-01-29"),  IdNum = "123456789123456789", Job = "木工", WorkDays = 10, IsSettled = false, IsCommited = true},
                new Departure{Order = 5,  Name = "吴亚亦", DepartureDate =  DateTime.Parse("2015-01-30"),  IdNum = "123456789123456789", Job = "木工", WorkDays = 10, IsSettled = true, IsCommited = false},
                new Departure{Order = 6,  Name = "王国祥", DepartureDate =  DateTime.Parse("2015-01-31"),  IdNum = "123456789123456789", Job = "木工", WorkDays = 10, IsSettled = false, IsCommited = true},
                new Departure{Order = 7,  Name = "马群东", DepartureDate =  DateTime.Parse("2015-02-01"),  IdNum = "123456789123456789", Job = "木工", WorkDays = 10, IsSettled = true, IsCommited = false},
                new Departure{Order = 8,  Name = "何  叶", DepartureDate =  DateTime.Parse("2015-02-01"),  IdNum = "123456789123456789", Job = "木工", WorkDays = 10, IsSettled = false, IsCommited = true},
                new Departure{Order = 9,  Name = "王  强", DepartureDate =  DateTime.Parse("2015-02-01"),  IdNum = "123456789123456789", Job = "木工", WorkDays = 10, IsSettled = true, IsCommited = false},
                new Departure{Order = 10,  Name = "范双元", DepartureDate =  DateTime.Parse("2015-02-01"),  IdNum = "123456789123456789", Job = "木工", WorkDays = 10, IsSettled = false, IsCommited = true},
            };

            var info = new Info
            {
                Basic = new Basic
                {
                    Company = "XXXX有限公司",
                    Project = "XXXX项目",
                    Team = "XXXX组",
                    StartDate = DateTime.Parse("2015-01-26"),
                    EndDate = DateTime.Parse("2015-02-01"),
                    EnterCount = 3,
                    DepartureCount = 2,
                    CurrentCount = 10
                },
                Enters = enters,
                Departures = departures
            };

            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "output.docx");

            Exporter.Export(info, filePath);
        }
    }
}
