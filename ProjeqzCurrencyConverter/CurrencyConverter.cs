using System;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Office.Interop.MSProject;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.SharePoint.Client;
using ProjeqzCurrencyConverter.ProjeqzWebService;
using Exception = System.Exception;
using ProjeqzCurrencyConverter.CurrencyServices;

namespace ProjeqzCurrencyConverter
{
    public partial class CurrencyConverter
    {
        public static string CustomFieldName = "Converted Cost";

        public string FromCurrencyCode = string.Empty;

        private void CurrencyConverterLoad(object sender, RibbonUIEventArgs e)
        {
            try
            {
                # region UnUsedCode
                /*
                // loading items here
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
                        if (dataRow["Currency"] != null && dataRow["CurrencyCode"] != null && !string.IsNullOrEmpty(dataRow["Currency"].ToString()) && !string.IsNullOrEmpty(dataRow["CurrencyCode"].ToString()))
                        {
                            MenuCurrencies.Items.Add(new RibbonButton());
                            ((RibbonButton)MenuCurrencies.Items.Last()).Label = @dataRow["Currency"] + " [" +
                                                                                 dataRow["CurrencyCode"] + "]";
                        }
                    }
                } */

                /*                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"AFA" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"ALL" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"DZD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"ADF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"AON" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"ARS" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"AWG" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"AUD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"ATS" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"BSD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"BHD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"BDT" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"BBD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"BEF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"BZD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"XOF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"BMD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"BTN" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"BOB" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"BWP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"BRC" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"BND" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"BGL" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"XOF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"BIF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"KHR" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"XAF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"CAD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"CVE" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"KYD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"XAF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"XAF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"CLP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"CNY" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"COP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"KMF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"XAF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"CRC" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"HRK" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"CUP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"CVP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"CSK" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"DKK" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"DJF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"DOP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"DOP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"ECS" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"EGP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"SVC" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"XAF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"EEK" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"ETB" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"FKP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"FJD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"FIM" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"FRF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"XAF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"GMD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"DEM" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"GHC" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"GIP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"GBP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"GRD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"GTQ" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"GNF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"GWP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"GYD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"HTG" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"HNL" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"HKD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"HUF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"ISK" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"INR" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"IDR" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"IRR" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"IQD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"IEP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"ILS" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"ITL" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"JMD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"JPY" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"JOD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"KZT" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"KES" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"KWD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"LAK" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"LVL" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"LBP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"LSL" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"LRD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"LYD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"LTL" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"LUF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"MOP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"MWK" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"MYR" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"MVR" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"XOF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"MTL" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"MRO" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"MUR" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"MXP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"MNT" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"MAD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"MZM" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"MMK" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"NLG" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"ANG" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"NZD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"NIO" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"XOF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"NGN" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"KPW" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"NOK" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"OMR" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"PKR" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"XPD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"PAB" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"PGK" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"PYG" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"PEN" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"PHP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"PLZ" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"PTE" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"QAR" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"ROL" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"RUB" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"WST" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"SAR" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"XOF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"SCR" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"SLL" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"SGD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"SKK" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"SIT" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"SBD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"SOS" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"ZAR" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"ESP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"LKR" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"SDD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"SRG" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"SZL" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"SEK" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"CHF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"SYP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"TWD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"TZS" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"THB" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"XOF" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"TOP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"TTD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"TND" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"TRL" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"UGS" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"UAG" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"AED" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"GBP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"USD" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"UYP" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"VUV" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"VEB" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"VND" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"YUN" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"ZMK" });
                                CmbCurrency.Items.Add(new RibbonDropDownItem { Label = @"ZWD" }); */

                #endregion

                var app = Globals.ThisAddIn.Application;

                // getting all currencies and moving into an array
                var currencies = Enum.GetNames(typeof(Currency));
                foreach (string currency in currencies)
                {
                    CmbCurrency.Items.Add(new RibbonDropDownItem { Label = currency });
                    if (app.ActiveProject.CurrencyCode == currency)
                        CmbCurrency.Text = currency;
                }

                if (Constants.Edition == ProductEdition.Standard)
                   TxtCurrencyRate.Visible = true;
                else
                {
                    TxtCurrencyRate.Visible = false;
                    FromCurrencyCode = app.ActiveProject.CurrencyCode;
                }

            }
            catch (Exception)
            {
                return;
            }
        }

        private void BtnConvertCurrencyClick(object sender, RibbonControlEventArgs e)
        {
            if (!string.IsNullOrEmpty(CmbCurrency.Text))
            {
                var app = Globals.ThisAddIn.Application;
                if (Constants.Edition == ProductEdition.Standard)
                {
                    if (!string.IsNullOrEmpty(TxtCurrencyRate.Text))
                    {
                        try
                        {
                            var flag = false;

                            // setting default currency settings to MS-Project
                            var currencySymbol = SetDefaultCurrencySettingsToMsProject(app);

                            //app.ActiveProject.CurrencyDigits = 2;
                            foreach (Task task in app.ActiveProject.Tasks)
                            {
                                flag = true;
                                Double cost = 0.0d;

                                // getting value from custom field for the current task
                                var value =
                                    task.GetField(app.FieldNameToFieldConstant(CustomFieldName, PjFieldType.pjTask));
                                try
                                {
                                    // converting the rate
                                    cost = double.Parse(value, NumberStyles.Currency)*Convert.ToDouble(TxtCurrencyRate.Text);
                                }
                                catch (Exception)
                                {
                                    // in case of exception , try to remove special symbol and apply the calculation
                                    cost = Convert.ToDouble(value.Replace(currencySymbol, ""))*Convert.ToDouble(TxtCurrencyRate.Text);
                                }

                                // finally setting the value to custom field for the currenct task
                                task.SetField(app.FieldNameToFieldConstant(CustomFieldName, PjFieldType.pjTask),cost.ToString());

                            }
                            if (flag)
                                ShowBallon(@"Converting process completed successfully.", ToolTipIcon.Info);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, @"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            ShowBallon(ex.Message, ToolTipIcon.Error);
                        }
                    }
                }
                else // enterprise edition code starts here
                {
                    
                    if (FromCurrencyCode != CmbCurrency.Text)
                    {
                        // initiating the currency web sevice client
                        var currency = new CurrencyConvertor();
                        // getting the currency from the current iterated element
                        var primaryCurrency = (Currency)Enum.Parse(typeof(Currency), FromCurrencyCode);

                        // getting the currency from the current iterated element
                        var secondaryCurrency = (Currency)Enum.Parse(typeof(Currency), CmbCurrency.Text);

                        var rate = 0.0d;

                        if(Constants.ConvertionRatePullMethod == ConvertionRatePullMethod.GetFromLiveWebService)
                            // calling web method to get actual convertion rate
                            rate = currency.ConversionRate(primaryCurrency, secondaryCurrency);

                        else if(Constants.ConvertionRatePullMethod == ConvertionRatePullMethod.GetFromIntranetPwaServer) // if it needs to call intranet web service for conversion
                        {
                            var webServiceUrl = app.ActiveProject.ServerURL + "/_layouts/SPProjeqzCurrencyConverter/ProjeqzCurrencyConverter.asmx";
                            var projeqzCurrencyConverterService = new ProjeqzCurrencyConverterService
                                                                      {
                                                                          Url = webServiceUrl,
                                                                          AllowAutoRedirect = true,
                                                                          UseDefaultCredentials = true
                                                                      };
                            rate = projeqzCurrencyConverterService.Convert(primaryCurrency.ToString(), secondaryCurrency.ToString());
                        }
                        else
                        {
                            // finally option, get the rate from project convertion setting list from pwa site using client object model

                            var context = new ClientContext(app.ActiveProject.ServerURL);

                            var spList = context.Site.RootWeb.Lists.GetByTitle(Constants.ListName);

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
						                                                    <FieldRef Name='" + XmlConvert.EncodeName(Constants.FromCurrencyFieldName) + @"' />
							                                                    <Value Type='Choice'>" + primaryCurrency.ToString() + @"</Value>
					                                                    </Eq>
					                                                    <Eq>
						                                                    <FieldRef Name='" + XmlConvert.EncodeName(Constants.ToCurrencyFieldName) + @"' />
							                                                    <Value Type='Choice'>" + secondaryCurrency.ToString() + @"</Value>
					                                                    </Eq>
				                                                    </And>
				                                                    <And>
					                                                    <Eq>
						                                                    <FieldRef Name='" + XmlConvert.EncodeName(Constants.FromCurrencyFieldName) + @"' />
							                                                    <Value Type='Choice'>" + secondaryCurrency.ToString() + @"</Value>
					                                                    </Eq>
                                                                        <Eq>
						                                                    <FieldRef Name='" + XmlConvert.EncodeName(Constants.ToCurrencyFieldName) + @"' />
							                                                    <Value Type='Choice'>" + primaryCurrency.ToString() + @"</Value>
					                                                    </Eq>
					                                                    
				                                                    </And>
			                                                    </Or>
		                                                    </Where>
	                                                    </Query>
                                                        <ViewFields>
                                                            <FieldRef Name='" + XmlConvert.EncodeName(Constants.RateFieldName) + @"' />
                                                        </ViewFields>
                                                    </View>
                                                    "
                                };


                                var spItemCollection = spList.GetItems(camlQuery);

                                context.Load(spItemCollection);

                                context.ExecuteQuery();

                                foreach (ListItem listItem in spItemCollection)
                                {
                                    rate = Convert.ToDouble(listItem[Constants.RateFieldName]);
                                    break;
                                }
                            }
                        }

                        // iterating to all the tasks to set the new conversion
                        bool flag = false;
                        foreach (Task task in app.ActiveProject.Tasks)
                        {
                            flag = true;
                            var value = task.GetField(app.FieldNameToFieldConstant(CustomFieldName, PjFieldType.pjTask)).Replace(app.ActiveProject.CurrencySymbol,"");
                            try
                            {
                                Double cost = double.Parse(value, NumberStyles.Currency) * rate;
                                task.SetField(app.FieldNameToFieldConstant(CustomFieldName, PjFieldType.pjTask), cost.ToString(CultureInfo.InvariantCulture));
                            }
                            catch (Exception)
                            {
                                continue;
                            }
                        }

                        // setting default convertion rate to Ms Project
                        SetDefaultCurrencySettingsToMsProject(app);

                        FromCurrencyCode = CmbCurrency.Text;

                        if (flag)
                            ShowBallon(@"Converting process completed successfully.", ToolTipIcon.Info);
                    }
                }
            }
        }

        public bool TryGetCurrencySymbol(string isoCurrencySymbol, out string symbol)
        {
            symbol = CultureInfo
                .GetCultures(CultureTypes.AllCultures)
                .Where(c => !c.IsNeutralCulture)
                .Select(culture =>
                {
                    try
                    {
                        return new RegionInfo(culture.LCID);
                    }
                    catch
                    {
                        return null;
                    }
                })
                .Where(ri => ri != null && ri.ISOCurrencySymbol == isoCurrencySymbol)
                .Select(ri => ri.CurrencySymbol)
                .FirstOrDefault();
            return symbol != null;
        }

        public void ShowBallon(string message,ToolTipIcon toolTipIcon)
        {
            NotificationIcon.BalloonTipIcon = toolTipIcon;
            NotificationIcon.BalloonTipText = message;
            NotificationIcon.Visible = true;
            NotificationIcon.ShowBalloonTip(10000);
        }

        public string SetDefaultCurrencySettingsToMsProject(Microsoft.Office.Interop.MSProject.Application app)
        {
            string currencySymbol = string.Empty;
            try
            {
                
                app.ActiveProject.CurrencyCode = CmbCurrency.Text;
                
                if (TryGetCurrencySymbol(CmbCurrency.Text, out currencySymbol))
                {
                    app.ActiveProject.CurrencySymbol = currencySymbol;
                    app.ActiveProject.CurrencySymbolPosition = PjPlacement.pjBeforeWithSpace;
                }
            }
            catch (Exception)
            {
                // need to figure it out the exact the way for this 
            }
            return currencySymbol;
        }
    }

    
}