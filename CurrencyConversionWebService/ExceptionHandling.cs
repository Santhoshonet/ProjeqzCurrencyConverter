
using System;
using Microsoft.SharePoint.Administration;

namespace CurrencyConversionWebService
{
    public class ExceptionHandling
    {
        public static void WriteUlsLog(Exception ex)
        {
            SPDiagnosticsService.Local.WriteTrace(0, new SPDiagnosticsCategory(Constants.UlsLogCategoryName, TraceSeverity.Monitorable, EventSeverity.Error), TraceSeverity.Monitorable, ex.Message, new object[] { ex.Message });
        }
    }
}