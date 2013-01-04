using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Xml;
using ConsoleApplication1.CurrencyService;
using ConsoleApplication1.CustomFieldsWebSvc;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            /*
            var customFields = new CustomFields();
            customFields.Credentials = new System.Net.NetworkCredential("Administrator", "password@123");
            foreach (CustomFieldDataSet.CustomFieldsRow customFieldsRow in customFields.ReadCustomFields(string.Empty, false).CustomFields)
            {
                if (customFieldsRow.MD_PROP_NAME == "Converted Cost")
                {
                    
                }
            } 
            Console.WriteLine("CULTURE ISO ISO WIN DISPLAYNAME     ENGLISHNAME");
            foreach (CultureInfo ci in CultureInfo.GetCultures(CultureTypes.AllCultures))
            {
                //Console.WriteLine("{0,-7}", ci.Name);
                //Console.WriteLine(" {0,-3}", ci.TwoLetterISOLanguageName);
                //  Console.WriteLine(" {0,-3}", ci.ThreeLetterISOLanguageName);
                //Console.WriteLine(" {0,-3}", ci.ThreeLetterWindowsLanguageName);
                //Console.WriteLine(" {0,-40}", ci.DisplayName);
                // Console.WriteLine(" {0,-40}", ci.EnglishName);
            }
            Console.ReadKey();
            return; */
            var country = new country();
            string outPut = country.GetCurrencies();
            var dataSet = new DataSet();

            var reader = new XmlTextReader(new StringReader(outPut));
            reader.Read();

            dataSet.ReadXml(reader);

            foreach (DataTable dataTable in dataSet.Tables)
            {
                foreach (DataRow dataRow in dataTable.Rows)
                {
                    //if (dataRow["CurrencyCode"] != null && !string.IsNullOrEmpty(dataRow["CurrencyCode"].ToString()))
                    //{
                    //    Console.Write("CmbCurrency.Items.Add(new RibbonDropDownItem() { Label = \"");
                    //    Console.Write(dataRow["CurrencyCode"]);
                    //    Console.Write("\" });");
                    //    Console.WriteLine("");
                    //}
                    foreach (DataColumn dataColumn in dataTable.Columns)
                    {
                        Console.WriteLine(dataColumn.ColumnName + " : " + dataRow[dataColumn]);
                    }
                }
            }
            Console.ReadKey();
        }
    }
}