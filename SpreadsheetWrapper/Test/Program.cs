using System;
using System.Collections.Generic;
using System.Data;
using SpreadsheetWrapper;
using System.Linq;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var spreadsheet1 = new Spreadsheet())
            {
                // ------------ PREPARE DATA ------------

                IEnumerable<Person> people = GetTestCollection();
                DataTable cars = GetTestDataTable();

                // ------------ INSERT ROW(S) / TABLE ------------

                // Insert all people
                var sheet1 = spreadsheet1.GetSheetByNameWithHeader("Everyone", new string[] { "Name", "Age", "IsMale", "Date of birth" });
                spreadsheet1.InsertRows(sheet1, people.Select(x => x.ConvertToRow()), autofit: false);

                // Insert one single row
                spreadsheet1.InsertRowAtEnd(sheet1, new Person { Name = "Dog", Age = 3, DateOfBirth = new DateTime(2018, 4, 21), IsMale = true }.ConvertToRow());
                spreadsheet1.Autofit(sheet1);

                // Insert adults only (new Adults sheet will be inserted), no header and will leave one blank column and one blank row
                spreadsheet1.InsertRows(sheetName: "Adults", people.Where(x => x.Age >= 18).Select(x => x.ConvertToRow()), rowIndex: 2, columnIndex: 2, autofit: true);

                // Insert cars, first using default row generator then using custom row generator
                spreadsheet1.InsertTable(cars, includeHeader: true);
                var sheet3 = spreadsheet1.GetSheetByName(cars.TableName);

                spreadsheet1.InsertTable(cars, sheetName: "custom built cars", includeHeader: true, rowConverter: (x) => Car.ConvertToRow(x), autofit: true, rowIndex: 10);
                spreadsheet1.InsertImage("custom built cars", "tux.png", imageHeightPixel: 75, imageWidthPixel: 125);

                // ------------ SAVE AS EXCEL / CSV / SERIALIZE DATA ------------

                // Excel with 4 sheets
                spreadsheet1.SaveExcelAs("test");
                spreadsheet1.SaveExcelAs("test2", password: "optional_password");

                // 4 Csv files, will create a new folder and place the files inside, one per sheet
                spreadsheet1.SaveCsvAs(basePath: "temp");

                // 1 Csv file, for sheet1
                spreadsheet1.SaveCsvAs(sheet1, "test");

                // Serialize "would be" content of Excel in byte[] form
                var bytes1 = spreadsheet1.SerializeExcel();

                // Serialize "would be" content of csv sheet3 in byte[] form
                var bytes2 = spreadsheet1.SerializeCsv(sheet3);

                // String content of csv sheet3
                List<string> csvSheet3 = spreadsheet1.GetSheetAsCsv(sheet3);
                string allInOne = spreadsheet1.GetSheetAsCsvString(sheet3);
            }
        }

        private static DataTable GetTestDataTable()
        {
            var dt = new DataTable("Cars");
            dt.Columns.Add("Maker", typeof(string));
            dt.Columns.Add("NumerOfPassengers", typeof(int));
            dt.Columns.Add("TurboMode", typeof(bool));
            dt.Columns.Add("Year", typeof(DateTime));
            dt.Rows.Add(new object[] { "Honda", 5, false, new DateTime(1990, 7, 3) });
            dt.Rows.Add(new object[] { "Mazda", 2, true, new DateTime(1986, 1, 30) });
            dt.Rows.Add(new object[] { "Toyota", 5, false, new DateTime(1950, 8, 18) });
            dt.Rows.Add(new object[] { "Ford", 7, false, new DateTime(1955, 8, 11) });
            return dt;
        }

        private static IEnumerable<Person> GetTestCollection()
        {
            var list = new List<Person>();
            list.AddRange(new[] {
                new Person { Name = "Bart", Age = 30, IsMale = true, DateOfBirth = new DateTime(1990, 7, 3) },
                new Person { Name = "Lisa", Age = 16, IsMale = false, DateOfBirth = new DateTime(1986, 1, 30) },
                new Person { Name = "Homer", Age = 70, IsMale = true, DateOfBirth = new DateTime(1950, 8, 18) },
                new Person { Name = "Marge", Age =65, IsMale = false, DateOfBirth = new DateTime(1955, 8, 11) }
            });
            return list;
        }

        private class Person
        {
            public string Name { get; set; }
            public int Age { get; set; }
            public bool IsMale { get; set; }
            public DateTime DateOfBirth { get; set; }

            public IEnumerable<string> ConvertToRow()
            {
                return new List<string>
                {
                    Name,
                    Age.ToString(),
                    IsMale ? "Yes" : "No",
                    DateOfBirth.ToString("yyyy-MM-dd")
                };
            }
        }

        private class Car
        {
            public string Maker { get; set; }
            public DateTime Year { get; set; }
            public int NumberOfPassengers { get; set; }
            public bool TurboMode { get; set; }

            public static IEnumerable<string> ConvertToRow(object[] car)
            {
                return new string[]
                {
                    car[0].ToString(),
                    car[1].ToString(),
                    Convert.ToBoolean(car[2]) ? "Yes" : "No",
                    Convert.ToDateTime(car[3]).ToString("yyyy-MM-dd")
                };
            }
        }
    }
}
