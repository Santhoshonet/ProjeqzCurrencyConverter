using System;
using System.Web.Services;
using CurrencyConversionWebService.CurrencyServices;

namespace CurrencyConversionWebService
{
    [WebService(Namespace = "http://microsoft.com/webservices/")]
    public class ProjeqzCurrencyConverterService : WebService
    {
        [WebMethod]
        public double Convert(string fromCurrency, string toCurrency)
        {
            try
            {
                // initiating the currency web sevice client
                var currency = new CurrencyConvertor();

                // getting the currency from the current iterated element
                var primaryCurrency = (Currency)Enum.Parse(typeof(Currency), fromCurrency);

                // getting the currency from the current iterated element
                var secondaryCurrency = (Currency)Enum.Parse(typeof(Currency), toCurrency);


                // calling web method to get actual convertion rate
                return currency.ConversionRate(primaryCurrency, secondaryCurrency);
            }
            catch (Exception ex)
            {
               ExceptionHandling.WriteUlsLog(ex);
               return 0.0f;
            }
        }
    }
}