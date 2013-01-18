using System;
using System.Data;
using System.IO;
using System.Xml;
using ConsoleApplication1.CurrencyService;
using ConsoleApplication1.CurrencyServices;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Utilities;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            // declaring all contants and names 
            const string siteUrl = "http://sp2010demo/pwa/";
            const string listName = "ProjeqzConversionSettings";
            const string listDescription = "List to keep all currency settings";
            const string fromCurrencyFieldName =  "From Currency";
            const string toCurrencyFieldName = "To Currency";
            const string rateFieldName = "Rate";

            string fromCurrencyValue = "ALL";
            string toCurrencyValue = "DZD";
            // pulling value from client object model

            var context = new ClientContext(siteUrl);

            var spList = context.Site.RootWeb.Lists.GetByTitle(listName);

            context.Load(spList);

            context.ExecuteQuery();

            if (spList != null)
            {
                var camlQuery = new CamlQuery
                {
                    ViewXml = @"                                                           
                                                        <View>
	                                                    <Query>
		                                                    <Where>
			                                                    <Or>
				                                                    <And>
					                                                    <Eq>
						                                                    <FieldRef Name='" + XmlConvert.EncodeName(fromCurrencyFieldName) + @"' />
							                                                    <Value Type='Choice'>" + fromCurrencyValue + @"</Value>
					                                                    </Eq>
					                                                    <Eq>
						                                                    <FieldRef Name='" + XmlConvert.EncodeName(toCurrencyFieldName) + @"' />
							                                                    <Value Type='Choice'>" + toCurrencyValue + @"</Value>
					                                                    </Eq>
				                                                    </And>
				                                                    <And>
					                                                    <Eq>
						                                                    <FieldRef Name='" + XmlConvert.EncodeName(fromCurrencyFieldName) + @"' />
							                                                    <Value Type='Choice'>" + toCurrencyValue + @"</Value>
					                                                    </Eq>
                                                                        <Eq>
						                                                    <FieldRef Name='" + XmlConvert.EncodeName(toCurrencyFieldName) + @"' />
							                                                    <Value Type='Choice'>" + fromCurrencyValue + @"</Value>
					                                                    </Eq>
					                                                    
				                                                    </And>
			                                                    </Or>
		                                                    </Where>
	                                                    </Query>
                                                        <ViewFields>
                                                            <FieldRef Name='" + XmlConvert.EncodeName(rateFieldName) + @"' />
                                                        </ViewFields>
                                                    </View>
                                                    "
                };


                var spItemCollection = spList.GetItems(camlQuery);

                context.Load(spItemCollection);

                context.ExecuteQuery();

                foreach (ListItem listItem in spItemCollection)
                {
                   Console.WriteLine(listItem[rateFieldName].ToString());   
                }
            }

           // return;

            // impersonation for sharepoint 
            SPSecurity.RunWithElevatedPrivileges(delegate
            {
                // getting all currencies and moving into an array
                var currencies = Enum.GetNames(typeof(Currency));

                // creating SPSite object to access site values
                using (var site = new SPSite(siteUrl))
                {
                    // opening root site and moving to web object
                                                                 foreach (SPWeb web in site.AllWebs)
                                                                 {

                                                                     // as a work around we are finding pwa with title, need to figure our an elagant way for this
                                                                     if (web.Title == "Project Web App")
                                                                     {
                                                                         // setting allow unsafe updates to true for avoiding list item creating issues.
                                                                         web.AllowUnsafeUpdates = true;

                                                                         // getting list object here with name
                                                                         var list =
                                                                             site.RootWeb.Lists.TryGetList(listName);

                                                                         // checks if the list is found or not
                                                                         if (list == null)
                                                                         {
                                                                             // creating list if it is not exists
                                                                             var listGuid =
                                                                                 site.RootWeb.Lists.Add(listName,
                                                                                                        listDescription,
                                                                                                        SPListTemplateType
                                                                                                            .GenericList);

                                                                             // getting newly created list into object
                                                                             list = site.RootWeb.Lists[listGuid];

                                                                             // adding choice field to store from currency name
                                                                             list.Fields.Add(fromCurrencyFieldName,
                                                                                             SPFieldType.Choice, true);

                                                                             // getting newly created from currency field into object
                                                                             var spFromCurrencyField =
                                                                                 (SPFieldChoice)
                                                                                 list.Fields[fromCurrencyFieldName];

                                                                             // adding choice field to store 'to currency' name
                                                                             list.Fields.Add(toCurrencyFieldName,
                                                                                             SPFieldType.Choice, true);

                                                                             // getting newly created 'to currency' field into object
                                                                             var spToCurrencyField =
                                                                                 (SPFieldChoice)
                                                                                 list.Fields[toCurrencyFieldName];

                                                                             // iterating all currency types and move into choices in choice field
                                                                             foreach (string currencyName in currencies)
                                                                             {
                                                                                 // first adding to from currency field
                                                                                 spFromCurrencyField.Choices.Add(
                                                                                     currencyName);

                                                                                 // adding to to currency field
                                                                                 spToCurrencyField.Choices.Add(
                                                                                     currencyName);
                                                                             }

                                                                             // updating from currency field to store the updated/added changes
                                                                             spFromCurrencyField.Update();

                                                                             // updating 'to currency' field to store the updated/added changes
                                                                             spToCurrencyField.Update();

                                                                             // adding number field to store currency Rate
                                                                             list.Fields.Add(rateFieldName,
                                                                                             SPFieldType.Number, true);

                                                                             // getting newly created number field and moving into an object
                                                                             var rateField =
                                                                                 (SPFieldNumber)
                                                                                 list.Fields[rateFieldName];

                                                                             // setting number of decimals here
                                                                             rateField.DisplayFormat =
                                                                                 SPNumberFormatTypes.FourDecimals;

                                                                             // setting minimum value here
                                                                             rateField.MinimumValue = 0.0f;

                                                                             // setting default value as 0 
                                                                             rateField.DefaultValue = "0";

                                                                             // updating number filed to store updated/added values.
                                                                             rateField.Update();


                                                                             // updating the default view to show the new fields
                                                                             for (int viewIndex = 0;
                                                                                  viewIndex < list.Views.Count;
                                                                                  viewIndex++)
                                                                             {
                                                                                 SPView spView = list.Views[viewIndex];
                                                                                 // first removing all the fields
                                                                                 foreach (
                                                                                     string field in spView.ViewFields)
                                                                                 {
                                                                                     spView.ViewFields.Delete(field);
                                                                                 }

                                                                                 spView.Update();

                                                                                 // adding new fields
                                                                                 spView.ViewFields.Add(
                                                                                     spFromCurrencyField);
                                                                                 spView.ViewFields.Add(spToCurrencyField);
                                                                                 spView.ViewFields.Add(rateField);

                                                                                 spView.Update();
                                                                             }



                                                                             // finally updating the list to store above fields and its values, configurations.
                                                                             list.Update();
                                                                         }

                                                                         // initiating the currency web sevice client
                                                                         var currency = new CurrencyConvertor();

                                                                         // primary iteration for all the currencies
                                                                         for (var primaryIndex = 0;
                                                                              primaryIndex < currencies.Length;
                                                                              primaryIndex++)
                                                                         {

                                                                             // getting the currency from the current iterated element
                                                                             var primaryCurrency =
                                                                                 (Currency)
                                                                                 Enum.Parse(typeof (Currency),
                                                                                            currencies[primaryIndex]);

                                                                             // secondary iteration for rest of the currencies
                                                                             for (var secondaryIndex = primaryIndex + 1;
                                                                                  secondaryIndex < currencies.Length;
                                                                                  secondaryIndex++)
                                                                             {
                                                                                 // getting the currency from the current iterated element
                                                                                 var secondaryCurrency =
                                                                                     (Currency)
                                                                                     Enum.Parse(typeof (Currency),
                                                                                                currencies[
                                                                                                    secondaryIndex]);

                                                                                 // calling web service method to get today's currency rate and storing into variable
                                                                                 var rate =
                                                                                     currency.ConversionRate(
                                                                                         primaryCurrency,
                                                                                         secondaryCurrency);

                                                                                 // checking whether it is having any conversion rate for the given currencies
                                                                                 if (rate > 0)
                                                                                 {
                                                                                     // adding a list item to store the values
                                                                                     var newListItem = list.Items.Add();

                                                                                     // updating primary currency field value into from currency field
                                                                                     newListItem[fromCurrencyFieldName]
                                                                                         = primaryCurrency.ToString();

                                                                                     // updating secondary currency field value into 'to currency' field
                                                                                     newListItem[toCurrencyFieldName] =
                                                                                         secondaryCurrency.ToString();

                                                                                     // updating rate field with the real time value
                                                                                     newListItem[rateFieldName] = rate;

                                                                                     // finally update the new list item store in the content db
                                                                                     newListItem.Update();
                                                                                 }
                                                                             }
                                                                         }
                                                                     }
                                                                 }
                }
            }
        );


            return;
            
            


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