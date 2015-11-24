namespace ReDistr
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.buttonGetOrder = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.buttonGetTransfers = this.Factory.CreateRibbonButton();
            this.buttonMakeTransfersBook = this.Factory.CreateRibbonButton();
            this.buttonGetOrdersLists = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.labelPeriodSellingTo = this.Factory.CreateRibbonLabel();
            this.labelPeriodSelling = this.Factory.CreateRibbonLabel();
            this.labelStockDate = this.Factory.CreateRibbonLabel();
            this.labelPeriodSellingFrom = this.Factory.CreateRibbonLabel();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "ReDistr";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.buttonGetOrder);
            this.group1.Items.Add(this.buttonGetOrdersLists);
            this.group1.Label = "Заказы";
            this.group1.Name = "group1";
            // 
            // buttonGetOrder
            // 
            this.buttonGetOrder.Label = "Рассчитать заказы";
            this.buttonGetOrder.Name = "buttonGetOrder";
            this.buttonGetOrder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGetOrders_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.buttonGetTransfers);
            this.group2.Items.Add(this.buttonMakeTransfersBook);
            this.group2.Label = "Перемещения";
            this.group2.Name = "group2";
            // 
            // buttonGetTransfers
            // 
            this.buttonGetTransfers.Label = "Рассчитать";
            this.buttonGetTransfers.Name = "buttonGetTransfers";
            this.buttonGetTransfers.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGetTransfers_Click);
            // 
            // buttonMakeTransfersBook
            // 
            this.buttonMakeTransfersBook.Label = "Сформировать файлы";
            this.buttonMakeTransfersBook.Name = "buttonMakeTransfersBook";
            this.buttonMakeTransfersBook.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonMakeTransfersBook_Click);
            // 
            // buttonGetOrdersLists
            // 
            this.buttonGetOrdersLists.Label = "Рассчитать списки заказов";
            this.buttonGetOrdersLists.Name = "buttonGetOrdersLists";
            this.buttonGetOrdersLists.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGetOrderLists_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.label1);
            this.group3.Items.Add(this.labelPeriodSellingFrom);
            this.group3.Items.Add(this.labelPeriodSellingTo);
            this.group3.Items.Add(this.labelPeriodSelling);
            this.group3.Items.Add(this.separator1);
            this.group3.Items.Add(this.label2);
            this.group3.Items.Add(this.labelStockDate);
            this.group3.Label = "Информация";
            this.group3.Name = "group3";
            // 
            // label1
            // 
            this.label1.Label = "Отчет \"Продажи\"";
            this.label1.Name = "label1";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // label2
            // 
            this.label2.Label = "Отчет \"Остатки\"";
            this.label2.Name = "label2";
            // 
            // labelPeriodSellingTo
            // 
            this.labelPeriodSellingTo.Label = "To";
            this.labelPeriodSellingTo.Name = "labelPeriodSellingTo";
            // 
            // labelPeriodSelling
            // 
            this.labelPeriodSelling.Label = "count";
            this.labelPeriodSelling.Name = "labelPeriodSelling";
            // 
            // labelStockDate
            // 
            this.labelStockDate.Label = "StockDate";
            this.labelStockDate.Name = "labelStockDate";
            // 
            // labelPeriodSellingFrom
            // 
            this.labelPeriodSellingFrom.Label = "from";
            this.labelPeriodSellingFrom.Name = "labelPeriodSellingFrom";
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGetOrder;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGetTransfers;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonMakeTransfersBook;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGetOrdersLists;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelPeriodSellingFrom;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelPeriodSellingTo;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelPeriodSelling;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelStockDate;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
