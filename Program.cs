using ExcelDataReader;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Utils.Model;

namespace Utils
{
    internal class Program
    {
        public static string beginTransaction = new string("BEGIN TRY\n" +
                     "  BEGIN TRANSACTION\n" +
                     //"      DELETE profilesection.countries\n" +
                     "      DECLARE @username as nvarchar(max);\n" +
                     "      SET @username = (SELECT USER_NAME());\n");

        public static string endTransaction = new string("      SET NOEXEC OFF\n" +
                     "  COMMIT\n" +
                     "END TRY\n" +
                     "BEGIN CATCH\n" +
                     "  SELECT ERROR_MESSAGE(), ERROR_LINE()\n" +
                     "    ROLLBACK\n" +
                     "END CATCH");

        public static string sqlCountryCommand = new string("INSERT profilesection.Countries(country_id, name, country_short_code, postal_code_validation_rule, " +
            "needs_state, prefix_regex, created_at, created_by, updated_at, updated_by) VALUES (N'{0}', N'{1}', N'{2}', N'{3}', N'{4}', N'{5}', " +
            "GETDATE(), @username, GETDATE(), @username)");

        public static string sqlStateCommand = new string("INSERT INTO profilesection.country_states (country_id, state, state_code_1, state_code_2, " +
            "state_code_3, state_code_4, state_code_5, prefix, created_at, created_by, updated_at, updated_by) VALUES (N'{0}', N'{1}', N'{2}', null, null, null" +
            ", null, N'{3}', GETDATE(), @username, GETDATE(), @username)");

        public static SortedList<string, string> ret = new SortedList<string, string>();

        public static List<CountryFileStructure> newFormat = new List<CountryFileStructure>();

        public static List<CountryStructure> json = new List<CountryStructure>();

        public static List<CountryStructure> countries = new List<CountryStructure>();

        public static List<CountryState> countriesState = new List<CountryState>();

        public static List<string> sqlcommands = new List<string>();

        private static void Main(string[] args)
        {
            
            CreateCountryScript();
            CreateStatesScript();
            CreateEUCountriesJSON();
        }

        public static void CreateCountryScript()
        {
            GetCountryName();
            UpdateCountriesName();
            SqlCountryScript();
        }

        public static void SqlCountryScript()
        {
            string prefix = "{\"Regex\":\".* \"}";
            foreach (CountryStructure country in countries)
            {
                switch (country.CountryShortCode)
                {
                    case "US":
                    case "PR":
                    case "IN":
                        country.PrefixRegex = "{\"Regex\":\"^\\d{3}\"}";
                        break;
                    case "TH":
                    case "JP":
                    case "MX":
                        country.PrefixRegex = "{\"Regex\":\"^\\d{2}\"}";
                        break;
                    case "CA":
                        country.PrefixRegex = "{\"Regex\":\"^(?:[ABCEGHJ-NPRSTVXY])\"}";
                        break;
                    default:
                        country.PrefixRegex = prefix;
                        break;

                }
                country.PrefixRegex = country.PrefixRegex.Replace("\\", "\\\\");
                country.PostalCodeValidationRule = "{\"Regex\":\"" + country.PostalCodeValidationRule + "\"} ";
                country.PostalCodeValidationRule = country.PostalCodeValidationRule.Replace("\\", "\\\\");

                sqlcommands.Add(String.Format(sqlCountryCommand, country.CountryGuid, country.Name, country.CountryShortCode, country.PostalCodeValidationRule, country.NeedsState, country.PrefixRegex));
            }

            //write to file
            TextWriter tw = new StreamWriter("CountrySQLScript.sql");
            tw.Write(beginTransaction);
            foreach (string s in sqlcommands)
            {
                tw.Write(s);
                tw.Write("\n");
            }
            tw.Write(endTransaction);
            tw.Close();
        }

        public static List<CountryStructure> GetCountryName()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = File.Open(@"Resources/CountriesforParcelFedex.xlsx", FileMode.Open, FileAccess.Read))
            {
                using (var streamGuids = File.Open(@"Resources/CountryGuids.xlsx", FileMode.Open, FileAccess.Read))
                {
                    // Auto-detect format, supports:
                    //  - Binary Excel files (2.0-2003 format; *.xls)
                    //  - OpenXml Excel files (2007 format; *.xlsx)
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        using (var readerGuids = ExcelReaderFactory.CreateReader(streamGuids))
                        {
                            // 2. Use the AsDataSet extension method
                            DataSet result = reader.AsDataSet();

                            DataSet resultGuids = readerGuids.AsDataSet();

                            DataTable data_table = result.Tables[0];

                            DataTable data_table_Guids = resultGuids.Tables[0];

                            for (int i = 1; i < data_table.Rows.Count; i++)
                            {
                                CountryStructure countryToInsert = new CountryStructure
                                {
                                    Name = data_table.Rows[i][0].ToString(),
                                    CountryShortCode = data_table.Rows[i][2].ToString(),
                                    CountryGuid = data_table_Guids.Rows[i-1][0].ToString()
                                };
                                if (data_table.Rows[i][5].ToString().Contains("YES"))
                                    countryToInsert.NeedsState = "1";
                                else
                                    countryToInsert.NeedsState = "0";

                                countryToInsert.IsEU = data_table.Rows[i][6].ToString().ToLower().Contains("YES".ToLower());

                                countries.Add(countryToInsert);
                            }

                            return countries;
                        }
                            
                    }
                }
                
            }
        }

        public static void UpdateCountriesName()
        {
            JObject o1 = JObject.Parse(File.ReadAllText("Resources/CountryZipCodeRegex.json"));
            var CountryZip = (JArray)o1["ZipCodes"];
            newFormat = CountryZip.ToObject<List<CountryFileStructure>>();

            for (int i = 0; i < newFormat.Count; i++)
            {
                for (int y = 0; y < countries.Count; y++)
                {
                    if (countries[y].CountryShortCode == newFormat[i].ISO)
                    {
                        if (String.IsNullOrEmpty(newFormat[i].Regex))
                            countries[y].PostalCodeValidationRule = ".*";
                        else
                            countries[y].PostalCodeValidationRule = newFormat[i].Regex;
                    }
                }
            }
        }

        public static void CreateStatesScript()
        {
            ReadInfo();
            SqlStatScript();
        }

        public static void ReadInfo()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = File.Open(@"Resources/CountriesforParcelFedex.xlsx", FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // 2. Use the AsDataSet extension method
                    DataSet result = reader.AsDataSet();

                    //Excel Pages
                    DataTable data_table = result.Tables[2];

                    for (int i = 1; i < data_table.Rows.Count; i++)
                    {
                        CountryState countryStateToInsert = new CountryState
                        {
                            Country = data_table.Rows[i][0].ToString(),
                            Prefix = data_table.Rows[i][1].ToString(),
                            StateCode = data_table.Rows[i][2].ToString(),
                            StateName = data_table.Rows[i][3].ToString(),
                        };
                        countriesState.Add(countryStateToInsert);
                    }

                    /* TextWriter tw = new StreamWriter("CountryStateInfo.txt");
                     foreach (CountryState s in countriesState)
                     {
                         tw.Write(s.Country + " " + s.StateName + " " + s.StateCode + " " + s.Prefix);
                         tw.Write("\n");
                     }
                     tw.Close();*/
                }
            }
        }

        public static void SqlStatScript()
        {
            //Puerto Rico also has US as code
            string USID = countries.Where(c => c.Name == "U.S.A.").Select(c => c.CountryGuid).FirstOrDefault();
            string PRID = countries.Where(c => c.Name == "Puerto Rico").Select(c => c.CountryGuid).FirstOrDefault();
            string CNID = countries.Where(c => c.CountryShortCode == "CN").Select(c => c.CountryGuid).FirstOrDefault();
            string MXID = countries.Where(c => c.CountryShortCode == "MX").Select(c => c.CountryGuid).FirstOrDefault();
            string THID = countries.Where(c => c.CountryShortCode == "TH").Select(c => c.CountryGuid).FirstOrDefault();
            string INID = countries.Where(c => c.CountryShortCode == "IN").Select(c => c.CountryGuid).FirstOrDefault();
            string JPID = countries.Where(c => c.CountryShortCode == "JP").Select(c => c.CountryGuid).FirstOrDefault();
            string CAID = countries.Where(c => c.CountryShortCode == "CA").Select(c => c.CountryGuid).FirstOrDefault();
            foreach (CountryState country in countriesState)
            {
                if (String.IsNullOrEmpty(country.StateCode))
                    country.StateCode = "null";
                if (country.Country == "US")
                    sqlcommands.Add(String.Format(sqlStateCommand, USID, country.StateName, country.StateCode, country.Prefix));
                if (country.Country == "PR")
                    sqlcommands.Add(String.Format(sqlStateCommand, PRID, country.StateName, country.StateCode, country.Prefix));
                if (country.Country == "CN")
                    sqlcommands.Add(String.Format(sqlStateCommand, CNID, country.StateName, country.StateCode, country.Prefix));
                if (country.Country == "MX")
                    sqlcommands.Add(String.Format(sqlStateCommand, MXID, country.StateName, country.StateCode, country.Prefix));
                if (country.Country == "TH")
                    sqlcommands.Add(String.Format(sqlStateCommand, THID, country.StateName, country.StateCode, country.Prefix));
                if (country.Country == "IN")
                    sqlcommands.Add(String.Format(sqlStateCommand, INID, country.StateName, country.StateCode, country.Prefix));
                if (country.Country == "JP")
                    sqlcommands.Add(String.Format(sqlStateCommand, JPID, country.StateName, country.StateCode, country.Prefix));
                if (country.Country == "CA")
                    sqlcommands.Add(String.Format(sqlStateCommand, CAID, country.StateName, country.StateCode, country.Prefix));
            }

            //write to file
            TextWriter tw = new StreamWriter("StateSQLScript.sql");
            tw.Write(beginTransaction);
            foreach (string s in sqlcommands)
            {
                tw.Write(s);
                tw.Write("\n");
            }
            tw.Write(endTransaction);
            tw.Close();
        }

        private static void CreateEUCountriesJSON ()
        {
            EUCountriesModel euCountriesModel = new EUCountriesModel() { EU = new List<object>() };

            foreach (CountryStructure country in countries)
            {
                if (country.IsEU)
                {
                    euCountriesModel.EU.Add(new { country = country.CountryGuid.ToString() });
                }
            }

            string json = JsonConvert.SerializeObject(euCountriesModel, Formatting.Indented);

            using (var tw = new StreamWriter("EU-Countries.json", true))
            {
                tw.WriteLine(json.ToString());
                tw.Close();
            }
        }
    }
}