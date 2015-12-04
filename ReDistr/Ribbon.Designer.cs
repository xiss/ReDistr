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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon));
			this.tab1 = this.Factory.CreateRibbonTab();
			this.groupData = this.Factory.CreateRibbonGroup();
			this.buttonParseData = this.Factory.CreateRibbonButton();
			this.button1 = this.Factory.CreateRibbonButton();
			this.separator2 = this.Factory.CreateRibbonSeparator();
			this.checkBoxIncludeSellings = this.Factory.CreateRibbonCheckBox();
			this.checkBoxIncludeAdditionalParameters = this.Factory.CreateRibbonCheckBox();
			this.checkBoxIncludeCompetitorsFromAP = this.Factory.CreateRibbonCheckBox();
			this.groupOrders = this.Factory.CreateRibbonGroup();
			this.buttonGetOrder = this.Factory.CreateRibbonButton();
			this.buttonGetOrdersLists = this.Factory.CreateRibbonButton();
			this.groupTransfers = this.Factory.CreateRibbonGroup();
			this.buttonGetTransfers = this.Factory.CreateRibbonButton();
			this.buttonMakeTransfersBook = this.Factory.CreateRibbonButton();
			this.groupInfo = this.Factory.CreateRibbonGroup();
			this.label1 = this.Factory.CreateRibbonLabel();
			this.labelPeriodSelling = this.Factory.CreateRibbonLabel();
			this.labelPeriodSellingCount = this.Factory.CreateRibbonLabel();
			this.separator1 = this.Factory.CreateRibbonSeparator();
			this.label2 = this.Factory.CreateRibbonLabel();
			this.labelStockDate = this.Factory.CreateRibbonLabel();
			this.groupRevaluations = this.Factory.CreateRibbonGroup();
			this.buttonGetRevaluations = this.Factory.CreateRibbonButton();
			this.button2 = this.Factory.CreateRibbonButton();
			this.button3 = this.Factory.CreateRibbonButton();
			this.tab1.SuspendLayout();
			this.groupData.SuspendLayout();
			this.groupOrders.SuspendLayout();
			this.groupTransfers.SuspendLayout();
			this.groupInfo.SuspendLayout();
			this.groupRevaluations.SuspendLayout();
			// 
			// tab1
			// 
			this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.tab1.Groups.Add(this.groupData);
			this.tab1.Groups.Add(this.groupInfo);
			this.tab1.Groups.Add(this.groupOrders);
			this.tab1.Groups.Add(this.groupRevaluations);
			this.tab1.Groups.Add(this.groupTransfers);
			this.tab1.Label = "ReDistr";
			this.tab1.Name = "tab1";
			// 
			// groupData
			// 
			this.groupData.Items.Add(this.buttonParseData);
			this.groupData.Items.Add(this.button1);
			this.groupData.Items.Add(this.separator2);
			this.groupData.Items.Add(this.checkBoxIncludeSellings);
			this.groupData.Items.Add(this.checkBoxIncludeAdditionalParameters);
			this.groupData.Items.Add(this.checkBoxIncludeCompetitorsFromAP);
			this.groupData.Label = "Данные";
			this.groupData.Name = "groupData";
			// 
			// buttonParseData
			// 
			this.buttonParseData.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.buttonParseData.Image = ((System.Drawing.Image)(resources.GetObject("buttonParseData.Image")));
			this.buttonParseData.Label = "Считать данные";
			this.buttonParseData.Name = "buttonParseData";
			this.buttonParseData.ShowImage = true;
			this.buttonParseData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonParseData_Click);
			// 
			// button1
			// 
			this.button1.Label = "button1";
			this.button1.Name = "button1";
			this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
			// 
			// separator2
			// 
			this.separator2.Name = "separator2";
			// 
			// checkBoxIncludeSellings
			// 
			this.checkBoxIncludeSellings.Checked = true;
			this.checkBoxIncludeSellings.Description = "Книга с продажами";
			this.checkBoxIncludeSellings.Label = "Продажи";
			this.checkBoxIncludeSellings.Name = "checkBoxIncludeSellings";
			// 
			// checkBoxIncludeAdditionalParameters
			// 
			this.checkBoxIncludeAdditionalParameters.Checked = true;
			this.checkBoxIncludeAdditionalParameters.Description = "Книга с доп параметрами";
			this.checkBoxIncludeAdditionalParameters.Label = "Доп. параметры";
			this.checkBoxIncludeAdditionalParameters.Name = "checkBoxIncludeAdditionalParameters";
			// 
			// checkBoxIncludeCompetitorsFromAP
			// 
			this.checkBoxIncludeCompetitorsFromAP.Description = "Книга с конкурентами из П+";
			this.checkBoxIncludeCompetitorsFromAP.Label = "Конкуренты П+";
			this.checkBoxIncludeCompetitorsFromAP.Name = "checkBoxIncludeCompetitorsFromAP";
			// 
			// groupOrders
			// 
			this.groupOrders.Items.Add(this.buttonGetOrder);
			this.groupOrders.Items.Add(this.buttonGetOrdersLists);
			this.groupOrders.Label = "Заказы";
			this.groupOrders.Name = "groupOrders";
			// 
			// buttonGetOrder
			// 
			this.buttonGetOrder.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.buttonGetOrder.Enabled = false;
			this.buttonGetOrder.Image = ((System.Drawing.Image)(resources.GetObject("buttonGetOrder.Image")));
			this.buttonGetOrder.Label = "Рассчитать заказы";
			this.buttonGetOrder.Name = "buttonGetOrder";
			this.buttonGetOrder.ShowImage = true;
			this.buttonGetOrder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGetOrders_Click);
			// 
			// buttonGetOrdersLists
			// 
			this.buttonGetOrdersLists.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.buttonGetOrdersLists.Enabled = false;
			this.buttonGetOrdersLists.Image = ((System.Drawing.Image)(resources.GetObject("buttonGetOrdersLists.Image")));
			this.buttonGetOrdersLists.Label = "Рассчитать списки заказов";
			this.buttonGetOrdersLists.Name = "buttonGetOrdersLists";
			this.buttonGetOrdersLists.ShowImage = true;
			this.buttonGetOrdersLists.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGetOrderLists_Click);
			// 
			// groupTransfers
			// 
			this.groupTransfers.Items.Add(this.buttonGetTransfers);
			this.groupTransfers.Items.Add(this.buttonMakeTransfersBook);
			this.groupTransfers.Label = "Перемещения";
			this.groupTransfers.Name = "groupTransfers";
			// 
			// buttonGetTransfers
			// 
			this.buttonGetTransfers.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.buttonGetTransfers.Enabled = false;
			this.buttonGetTransfers.Image = ((System.Drawing.Image)(resources.GetObject("buttonGetTransfers.Image")));
			this.buttonGetTransfers.Label = "Рассчитать";
			this.buttonGetTransfers.Name = "buttonGetTransfers";
			this.buttonGetTransfers.ShowImage = true;
			this.buttonGetTransfers.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGetTransfers_Click);
			// 
			// buttonMakeTransfersBook
			// 
			this.buttonMakeTransfersBook.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.buttonMakeTransfersBook.Image = ((System.Drawing.Image)(resources.GetObject("buttonMakeTransfersBook.Image")));
			this.buttonMakeTransfersBook.Label = "Сформировать файлы";
			this.buttonMakeTransfersBook.Name = "buttonMakeTransfersBook";
			this.buttonMakeTransfersBook.ShowImage = true;
			this.buttonMakeTransfersBook.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonMakeTransfersBook_Click);
			// 
			// groupInfo
			// 
			this.groupInfo.Items.Add(this.label1);
			this.groupInfo.Items.Add(this.labelPeriodSelling);
			this.groupInfo.Items.Add(this.labelPeriodSellingCount);
			this.groupInfo.Items.Add(this.separator1);
			this.groupInfo.Items.Add(this.label2);
			this.groupInfo.Items.Add(this.labelStockDate);
			this.groupInfo.Label = "Информация";
			this.groupInfo.Name = "groupInfo";
			// 
			// label1
			// 
			this.label1.Label = "Продажи";
			this.label1.Name = "label1";
			// 
			// labelPeriodSelling
			// 
			this.labelPeriodSelling.Label = " ";
			this.labelPeriodSelling.Name = "labelPeriodSelling";
			// 
			// labelPeriodSellingCount
			// 
			this.labelPeriodSellingCount.Label = " ";
			this.labelPeriodSellingCount.Name = "labelPeriodSellingCount";
			// 
			// separator1
			// 
			this.separator1.Name = "separator1";
			// 
			// label2
			// 
			this.label2.Label = "Остатки";
			this.label2.Name = "label2";
			// 
			// labelStockDate
			// 
			this.labelStockDate.Label = " ";
			this.labelStockDate.Name = "labelStockDate";
			// 
			// groupRevaluations
			// 
			this.groupRevaluations.Items.Add(this.buttonGetRevaluations);
			this.groupRevaluations.Items.Add(this.button3);
			this.groupRevaluations.Label = "Переоценка";
			this.groupRevaluations.Name = "groupRevaluations";
			// 
			// buttonGetRevaluations
			// 
			this.buttonGetRevaluations.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.buttonGetRevaluations.Enabled = false;
			this.buttonGetRevaluations.Image = ((System.Drawing.Image)(resources.GetObject("buttonGetRevaluations.Image")));
			this.buttonGetRevaluations.Label = "Рассчитать переоценку";
			this.buttonGetRevaluations.Name = "buttonGetRevaluations";
			this.buttonGetRevaluations.ShowImage = true;
			this.buttonGetRevaluations.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGetRevaluations_Click);
			// 
			// button2
			// 
			this.button2.Label = "button2";
			this.button2.Name = "button2";
			this.button2.ShowImage = true;
			// 
			// button3
			// 
			this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
			this.button3.Label = "button3";
			this.button3.Name = "button3";
			this.button3.ShowImage = true;
			// 
			// Ribbon
			// 
			this.Name = "Ribbon";
			// 
			// Ribbon.OfficeMenu
			// 
			this.OfficeMenu.Items.Add(this.button2);
			this.RibbonType = "Microsoft.Excel.Workbook";
			this.Tabs.Add(this.tab1);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
			this.tab1.ResumeLayout(false);
			this.tab1.PerformLayout();
			this.groupData.ResumeLayout(false);
			this.groupData.PerformLayout();
			this.groupOrders.ResumeLayout(false);
			this.groupOrders.PerformLayout();
			this.groupTransfers.ResumeLayout(false);
			this.groupTransfers.PerformLayout();
			this.groupInfo.ResumeLayout(false);
			this.groupInfo.PerformLayout();
			this.groupRevaluations.ResumeLayout(false);
			this.groupRevaluations.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupOrders;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGetOrder;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupTransfers;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGetTransfers;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonMakeTransfersBook;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGetOrdersLists;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelPeriodSelling;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelPeriodSellingCount;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelStockDate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonParseData;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxIncludeSellings;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxIncludeAdditionalParameters;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxIncludeCompetitorsFromAP;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupRevaluations;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGetRevaluations;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
