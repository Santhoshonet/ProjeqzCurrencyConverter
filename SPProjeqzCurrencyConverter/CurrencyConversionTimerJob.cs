using System;
using CurrencyConversionWebService;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using CurrencyConversionWebService.CurrencyServices;
namespace SPProjeqzCurrencyConverter
{
    class CurrencyConversionTimerJob : SPJobDefinition
    {
        public CurrencyConversionTimerJob()
        {
            
        }
        public CurrencyConversionTimerJob(string jobName, SPWebApplication webApplication)
            : base(jobName, webApplication,null,SPJobLockType.None)
        {

            Title = jobName;
        }

        public override void Execute(Guid targetInstanceId)
        {
            try
            {
                // impersonation for sharepoint 
                SPSecurity.RunWithElevatedPrivileges(delegate
                                                         {
                                                             // getting all currencies and moving into an array
                                                             var currencies = Enum.GetNames(typeof (Currency));

                                                             // iterating SPSite object to access site values
                                                             foreach (SPSite site in WebApplication.Sites)
                                                             {

                                                                 // opening root site and moving to web object
                                                                 using (var web = site.RootWeb)
                                                                 {

                                                                     // as a work around we are finding pwa with title, need to figure our an elagant way for this
                                                                     if (web.Title == "Project Web App")
                                                                     {

                                                                         // setting allow unsafe updates to true for avoiding list item creating issues.
                                                                         web.AllowUnsafeUpdates = true;

                                                                         // getting list object here with name
                                                                         var list =
                                                                             site.RootWeb.Lists.TryGetList(
                                                                                 Constants.ListName);

                                                                         // checks if the list is found or not
                                                                         if (list == null)
                                                                         {
                                                                             // creating list if it is not exists
                                                                             var listGuid =
                                                                                 site.RootWeb.Lists.Add(
                                                                                     Constants.ListName,
                                                                                     Constants.ListDescription,
                                                                                     SPListTemplateType.GenericList);

                                                                             // getting newly created list into object
                                                                             list = site.RootWeb.Lists[listGuid];

                                                                             // adding choice field to store from currency name
                                                                             list.Fields.Add(
                                                                                 Constants.FromCurrencyFieldName,
                                                                                 SPFieldType.Choice, true);

                                                                             // getting newly created from currency field into object
                                                                             var spFromCurrencyField =
                                                                                 (SPFieldChoice)
                                                                                 list.Fields[
                                                                                     Constants.FromCurrencyFieldName];

                                                                             // adding choice field to store 'to currency' name
                                                                             list.Fields.Add(
                                                                                 Constants.ToCurrencyFieldName,
                                                                                 SPFieldType.Choice, true);

                                                                             // getting newly created 'to currency' field into object
                                                                             var spToCurrencyField =
                                                                                 (SPFieldChoice)
                                                                                 list.Fields[
                                                                                     Constants.ToCurrencyFieldName];

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
                                                                             list.Fields.Add(Constants.RateFieldName,
                                                                                             SPFieldType.Number, true);

                                                                             // getting newly created number field and moving into an object
                                                                             var rateField =
                                                                                 (SPFieldNumber)
                                                                                 list.Fields[Constants.RateFieldName];

                                                                             // setting number of decimals here
                                                                             rateField.DisplayFormat =
                                                                                 SPNumberFormatTypes.FourDecimals;

                                                                             // setting minimum value here
                                                                             rateField.MinimumValue = 0.0f;

                                                                             // setting default value as 0 
                                                                             rateField.DefaultValue = "0";

                                                                             // updating number filed to store updated/added values.
                                                                             rateField.Update();

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
                                                                                                currencies[primaryIndex]);

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
                                                                                     newListItem[
                                                                                         Constants.FromCurrencyFieldName
                                                                                         ] = primaryCurrency.ToString();

                                                                                     // updating secondary currency field value into 'to currency' field
                                                                                     newListItem[
                                                                                         Constants.ToCurrencyFieldName]
                                                                                         = secondaryCurrency.ToString();

                                                                                     // updating rate field with the real time value
                                                                                     newListItem[Constants.RateFieldName
                                                                                         ] = rate;

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
            }
            catch (Exception ex)
            {
                ExceptionHandling.WriteUlsLog(ex);
            }
            base.Execute(targetInstanceId);
        }
    }
}
