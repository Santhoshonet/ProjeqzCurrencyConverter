using Microsoft.Office.Tools.Ribbon;

namespace ProjeqzCurrencyConverter
{
    partial class CurrencyConverter : Microsoft.Office.Tools.Ribbon.OfficeRibbon
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public CurrencyConverter()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.tab2 = new Microsoft.Office.Tools.Ribbon.RibbonTab();
            this.GrpConverter = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.CmbCurrency = new Microsoft.Office.Tools.Ribbon.RibbonComboBox();
            this.TxtCurrencyRate = new Microsoft.Office.Tools.Ribbon.RibbonEditBox();
            this.BtnConvertCurrency = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.NotificationIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.tab2.SuspendLayout();
            this.GrpConverter.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.GrpConverter);
            this.tab2.Label = "Currency Converter";
            this.tab2.Name = "tab2";
            // 
            // GrpConverter
            // 
            this.GrpConverter.Items.Add(this.CmbCurrency);
            this.GrpConverter.Items.Add(this.TxtCurrencyRate);
            this.GrpConverter.Items.Add(this.BtnConvertCurrency);
            this.GrpConverter.Label = "Projeqz Currency Converter";
            this.GrpConverter.Name = "GrpConverter";
            // 
            // CmbCurrency
            // 
            this.CmbCurrency.Label = "Select currency";
            this.CmbCurrency.Name = "CmbCurrency";
            this.CmbCurrency.Text = null;
            // 
            // TxtCurrencyRate
            // 
            this.TxtCurrencyRate.Label = "Input currency";
            this.TxtCurrencyRate.Name = "TxtCurrencyRate";
            this.TxtCurrencyRate.Text = null;
            // 
            // BtnConvertCurrency
            // 
            this.BtnConvertCurrency.Label = "                         Convert                         ";
            this.BtnConvertCurrency.Name = "BtnConvertCurrency";
            this.BtnConvertCurrency.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.Button1Click);
            // 
            // NotificationIcon
            // 
            this.NotificationIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            this.NotificationIcon.Text = "Projeqz Currency Converter";
            this.NotificationIcon.Visible = true;
            // 
            // CurrencyConverter
            // 
            this.Name = "CurrencyConverter";
            this.RibbonType = "Microsoft.Project.Project";
            this.Tabs.Add(this.tab2);
            this.Load += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs>(this.CurrencyConverterLoad);
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.GrpConverter.ResumeLayout(false);
            this.GrpConverter.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion Component Designer generated code

        internal RibbonTab tab2;
        internal RibbonGroup GrpConverter;
        internal RibbonComboBox CmbCurrency;
        internal RibbonEditBox TxtCurrencyRate;
        internal RibbonButton BtnConvertCurrency;
        public System.Windows.Forms.NotifyIcon NotificationIcon;
    }

    partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonReadOnlyCollection
    {
        internal CurrencyConverter CurrencyConverter
        {
            get { return this.GetRibbon<CurrencyConverter>(); }
        }
    }
}