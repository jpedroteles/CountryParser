using ExcelDataReader;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Utils.Model;

namespace Utils
{
    internal class Program
    {
        public static string beginTransaction = new string("BEGIN TRY\n" +
                     "  BEGIN TRANSACTION\n" +
                     "      DELETE profilesection.countries\n" +
                     "      DECLARE @username as nvarchar(max);\n" +
                     "      SET @username = (SELECT USER_NAME());\n");

        public static string endTransaction = new string("      SET NOEXEC OFF\n" +
                     "  COMMIT\n" +
                     "END TRY\n" +
                     "BEGIN CATCH\n" +
                     "  SELECT ERROR_MESSAGE(), ERROR_LINE()\n" +
                     "    ROLLBACK\n" +
                     "END CATCH");

        public static string sqlCountryCommand = new string("INSERT profilesection.Countries(name, country_short_code, postal_code_validation_rule, " +
            "needs_state, prefix_regex, created_at, created_by, updated_at, updated_by) VALUES (N'{0}', N'{1}', {2}, {3}, {4}, " +
            "GETDATE(), @username, GETDATE(), @username)");

        public static string sqlStateCommand = new string("INSERT INTO profilesection.country_states (country_id, state, state_code_1, state_code_2, " +
            "state_code_3, state_code_4, state_code_5, prefix, created_at, created_by, updated_at, updated_by) VALUES ({0}, {1}, {2}, null, null, null" +
            ", null, N{3}, GETDATE(), @username, GETDATE(), @username)");

        public static SortedList<string, string> ret = new SortedList<string, string>();

        public static List<CountryFileStructure> newFormat = new List<CountryFileStructure>();

        public static List<CountryStructure> json = new List<CountryStructure>();

        public static List<CountryStructure> countries = new List<CountryStructure>();

        public static List<CountryState> countriesState = new List<CountryState>();

        public static List<string> sqlcommands = new List<string>();

        private static void Main(string[] args)
        {
            CreateStatesScript();
            //CreateCountryScript();
        }

        public static void CreateCountryScript()
        {
            GetCountryName();
            UpdateCountriesName();
            SqlCountryScript();
        }

        public static void SqlCountryScript()
        {
            string prefix = "'N\"{\"Regex\":\".* \"}'";
            foreach (CountryStructure country in countries)
            {
                if (country.CountryShortCode == "US" || country.CountryShortCode == "PR" || country.CountryShortCode == "IN")
                    country.PrefixRegex = "N'{\"Regex\":\"^\\d{3}\"}'";
                if (country.CountryShortCode == "CH" || country.CountryShortCode == "JP" || country.CountryShortCode == "MX")
                    country.PrefixRegex = "N'{\"Regex\":\"^\\d{2}\"}'";
                if (country.CountryShortCode == "CA")
                    country.PrefixRegex = "N'{\"Regex\":\"^(?:[ABCEGHJ-NPRSTVXY])\"}'";
                else
                    country.PrefixRegex = prefix;
                country.PostalCodeValidationRule = "N'{\"Regex\":\"" + country.PostalCodeValidationRule + "\"}' ";
                country.PostalCodeValidationRule = country.PostalCodeValidationRule.Replace("\\", "\\\\");

                sqlcommands.Add(String.Format(sqlCountryCommand, country.Name, country.CountryShortCode, country.PostalCodeValidationRule, country.NeedsState, country.PrefixRegex));
            }

            //write to file
            TextWriter tw = new StreamWriter("CountrySQLScript.txt");
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
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // 2. Use the AsDataSet extension method
                    DataSet result = reader.AsDataSet();

                    DataTable data_table = result.Tables[0];

                    for (int i = 1; i < data_table.Rows.Count; i++)
                    {
                        CountryStructure countryToInsert = new CountryStructure
                        {
                            Name = data_table.Rows[i][0].ToString(),
                            CountryShortCode = data_table.Rows[i][2].ToString(),
                        };
                        if (data_table.Rows[i][5].ToString().Contains("YES"))
                            countryToInsert.NeedsState = "1";
                        else
                            countryToInsert.NeedsState = "0";
                        countries.Add(countryToInsert);
                    }

                    return countries;
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
                            StateName = data_table.Rows[i][1].ToString(),
                            StateCode = data_table.Rows[i][2].ToString(),
                            Prefix = data_table.Rows[i][3].ToString(),
                        };
                        countriesState.Add(countryStateToInsert);
                    }

                    TextWriter tw = new StreamWriter("CountryStateInfo.txt");
                    foreach (CountryState s in countriesState)
                    {
                        tw.Write(s.Country + " " + s.StateName + " " + s.StateCode + " " + s.Prefix);
                        tw.Write("\n");
                    }
                    tw.Close();
                }
            }
        }

        public static void SqlStatScript()
        {
            string USID = "BEF2708C-F1D1-41D9-A95D-8F5794750462";
            string PRID = "15FA019A-1F72-4104-95F3-32640CB67843";
            string CNID = "89FC8188-4693-4C39-AFAC-1BEB55C0524E";
            string MXID = "10F99FE2-35F3-47F1-8E01-2CC641EA912C";
            string THID = "BE2D24EE-41C3-4677-A211-5420E319217A";
            string INID = "0581DAAB-383F-42AC-A3E9-800745047593";
            string JPID = "61263884-EC54-4187-8A33-87755BC8B5FF";
            string CAID = "DD12EA71-71AA-44AB-A8E4-DCBB58576219";
            foreach (CountryState country in countriesState)
            {
                if (country.Country == "US")
                    sqlcommands.Add(String.Format(sqlStateCommand, USID, country.Country, country.StateName, country.StateCode, country.Prefix));
                if (country.Country == "PR")
                    sqlcommands.Add(String.Format(sqlStateCommand, PRID, country.Country, country.StateName, country.StateCode, country.Prefix));
                if (country.Country == "CN")
                    sqlcommands.Add(String.Format(sqlStateCommand, CNID, country.Country, country.StateName, country.StateCode, country.Prefix));
                if (country.Country == "MX")
                    sqlcommands.Add(String.Format(sqlStateCommand, MXID, country.Country, country.StateName, country.StateCode, country.Prefix));
                if (country.Country == "TH")
                    sqlcommands.Add(String.Format(sqlStateCommand, THID, country.Country, country.StateName, country.StateCode, country.Prefix));
                if (country.Country == "IN")
                    sqlcommands.Add(String.Format(sqlStateCommand, INID, country.Country, country.StateName, country.StateCode, country.Prefix));
                if (country.Country == "JP")
                    sqlcommands.Add(String.Format(sqlStateCommand, JPID, country.Country, country.StateName, country.StateCode, country.Prefix));
                if (country.Country == "CA")
                    sqlcommands.Add(String.Format(sqlStateCommand, CAID, country.Country, country.StateName, country.StateCode, country.Prefix));
            }

            //write to file
            TextWriter tw = new StreamWriter("StateSQLScript.txt");
            tw.Write(beginTransaction);
            foreach (string s in sqlcommands)
            {
                tw.Write(s);
                tw.Write("\n");
            }
            tw.Write(endTransaction);
            tw.Close();
        }
    }
}