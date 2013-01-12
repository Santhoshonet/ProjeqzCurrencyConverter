
namespace CurrencyConversionWebService
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
    }
}
