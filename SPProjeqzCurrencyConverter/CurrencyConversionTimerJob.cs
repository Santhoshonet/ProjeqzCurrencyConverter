using System;
using CurrencyConversionWebService;
using Microsoft.SharePoint.Administration;
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
                SPLibrary.CreateCurrencyConvertionSettingsList(WebApplication);
                
            }
            catch (Exception ex)
            {
                ExceptionHandling.WriteUlsLog(ex);
            }
            base.Execute(targetInstanceId);
        }
    }
}
