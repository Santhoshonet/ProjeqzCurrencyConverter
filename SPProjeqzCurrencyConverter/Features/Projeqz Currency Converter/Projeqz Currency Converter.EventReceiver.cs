using System;
using System.Runtime.InteropServices;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

namespace SPProjeqzCurrencyConverter.Features.Projeqz_Currency_Converter
{
    /// <summary>
    /// This class handles events raised during feature activation, deactivation, installation, uninstallation, and upgrade.
    /// </summary>
    /// <remarks>
    /// The GUID attached to this class may be used during packaging and should not be modified.
    /// </remarks>

    [Guid("9d889383-615b-4e54-a2e2-4c5fc7f919fa")]
    public class ProjeqzCurrencyConverterEventReceiver : SPFeatureReceiver
    {
        // Uncomment the method below to handle the event raised after a feature has been activated.

        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {
            var site = properties.Feature.Parent as SPSite;
            // make sure the job isn't already registered
            if (site != null)
            {
                foreach (SPJobDefinition job in site.WebApplication.JobDefinitions)
                {
                    if (job.Name == Constants.TimerJobName)
                        job.Delete();
                }
                // install the job
                var currencyConversionTimerJob = new CurrencyConversionTimerJob(Constants.TimerJobName, site.WebApplication);
                //To perform the task on daily basis
                var schedule = new SPDailySchedule {BeginHour = 0, BeginMinute = 0, BeginSecond = 0};
                currencyConversionTimerJob.Schedule = schedule;
                currencyConversionTimerJob.Update();
            }
        }


        // Uncomment the method below to handle the event raised before a feature is deactivated.

        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {
            var site = properties.Feature.Parent as SPSite;
            // make sure the job isn't already registered
            if (site != null)
            {
                foreach (SPJobDefinition job in site.WebApplication.JobDefinitions)
                {
                    if (job.Name == Constants.TimerJobName)
                        job.Delete();
                }
            }
        }


        // Uncomment the method below to handle the event raised after a feature has been installed.

        //public override void FeatureInstalled(SPFeatureReceiverProperties properties)
        //{
        //}


        // Uncomment the method below to handle the event raised before a feature is uninstalled.

        //public override void FeatureUninstalling(SPFeatureReceiverProperties properties)
        //{
        //}

        // Uncomment the method below to handle the event raised when a feature is upgrading.

        //public override void FeatureUpgrading(SPFeatureReceiverProperties properties, string upgradeActionName, System.Collections.Generic.IDictionary<string, string> parameters)
        //{
        //}
    }
}
