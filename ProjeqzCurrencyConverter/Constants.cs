
namespace ProjeqzCurrencyConverter
{
   public class Constants
    {
        // declaring all contants and names 
        public const string SiteUrl = "http://sp2010demo/pwa/";
        public const string ListName = "ProjeqzConversionSettings";
        public const string ListDescription = "List to keep all currency settings";
        public const string FromCurrencyFieldName = "From Currency";
        public const string ToCurrencyFieldName = "To Currency";
        public const string RateFieldName = "Rate";

        public const string TimerJobName = "Projeqz Currency Conversion";
        public const string UlsLogCategoryName = "ProjeqzCurrencyConversion";

       public static bool IsitInDevelopmentMode = true;

       public static ProductEdition Edition = ProductEdition.Enterprise;

       public static ConvertionRatePullMethod ConvertionRatePullMethod = ConvertionRatePullMethod.GetFromIntranetPwaServer;
    }

   public enum ProductEdition
   {
       Standard,
       Enterprise
   }

   public enum ConvertionRatePullMethod
   {
       GetFromLiveWebService, // from http://www.webservicex.net/CurrencyConvertor.asmx?WSDL
       GetFromIntranetPwaServer, // from our customer web service <<PWAURL>>/_layouts/SPProjeqzCurrencyConverter/ProjeqzCurrencyConverter.asmx
       GetFromPwaSiteList // from the currency custom list from pwa site
   }
}
