using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ReDistr
{
    public partial class Transfers
    {
        private void Лист3_Startup(object sender, System.EventArgs e)
        {
        }

        private void Лист3_Shutdown(object sender, System.EventArgs e)
        {
        }

		#region Код, созданный конструктором VSTO

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Лист3_Startup);
            this.Shutdown += new System.EventHandler(Лист3_Shutdown);
        }

        #endregion

		// Выводит на лист перемещения из списка перемещений сгруппированные по направлениям
	    public void FillList(List<Transfer> transfers)
	    {
			// Список возможных направлений перемещений
			var UnitedTransfers = ReDistr.GetPossibleTransfers(SimpleStockFactory.CurrentFactory.GetAllStocks()).ToList();

			foreach (var unitedTransfer in UnitedTransfers)
			{
				// Выбираем перемещения сгруппированные по направлению и объедененные по ЗЧ
				var transferList = transfers.Where(
					transfer => transfer.StockFrom == unitedTransfer.StockFrom && transfer.StockTo == unitedTransfer.StockTo)
					.GroupBy(transfer => transfer.Item).ToList();

			}
	    }
    }
}
