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
	public partial class Control
	{
		private void Лист1_Startup(object sender, System.EventArgs e)
		{
		}

		private void Лист1_Shutdown(object sender, System.EventArgs e)
		{
		}

		#region Код, созданный конструктором VSTO

		/// <summary>
		/// Обязательный метод для поддержки конструктора - не изменяйте
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		private void InternalStartup()
		{
			this.buttonGetMoving.Click += new System.EventHandler(this.buttonGetMoving_Click);
			this.Startup += new System.EventHandler(this.Лист1_Startup);
			this.Shutdown += new System.EventHandler(this.Лист1_Shutdown);

		}

		#endregion

		private void buttonGetMoving_Click(object sender, EventArgs e)
		{
			// Парсим данные из файлов
			var parser = new Parser(this);
			var items = parser.Parse();

			// Подготавливаем данные
			ReDistr.PrepareData(items);

			// Выводим таблицу для тестов
			Globals.Test.FillListStocks(items);

			// Формируем перемещения
			var transfers = new List<Transfer>();
			// для обеспечения одного комплекта
			transfers = ReDistr.GetTransfersFirstLvl(items, transfers);
			// для обеспечения мин. остатка
			transfers = ReDistr.GetTransfersSecondLvl(items, transfers);

			// Выводим перемещения на лист для тестов
			Globals.Test.FillListTransfers(transfers, items);

			// Выводим параметры отчетов
			Globals.Control.FillReportsParameters();

			// Создаем список возможных перемещений
			// var movings = ReDistr.GetPossibleTransfers(SimpleStockFactory.CurrentFactory.GetAllStocks());
		}

		// Выводит параметры отчетов на страницу управления
		public void FillReportsParameters()
		{
			var resultRange = new dynamic[4, 1];

			resultRange[0, 0] = Config.PeriodSellingFrom;
			resultRange[1, 0] = Config.PeriodSellingTo;
			resultRange[2, 0] = Config.SellingPeriod;
			resultRange[3, 0] = Config.StockDate;

			Range["G3:G6"].Value = resultRange;
		}
	}
}
