using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
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
#warning Удалить потом, для отладки
#if(DEBUG)
			buttonGetMoving_Click(new object(), new EventArgs());
			buttonMakeTransfers_Click(new object(), new EventArgs());
#endif
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
			this.buttonMakeTransfers.Click += new System.EventHandler(this.buttonMakeTransfers_Click);
			this.Startup += new System.EventHandler(this.Лист1_Startup);
			this.Shutdown += new System.EventHandler(this.Лист1_Shutdown);

		}

		#endregion

		private void buttonGetMoving_Click(object sender, EventArgs e)
		{
			// Парсим данные из файлов
			var parser = new Parser();
			var items = parser.Parse();

			// Если парсинг не удался, выходим
			if (items == null)
			{
				return;
			}

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
			// для обеспечения необходимого запаса, перемещения создаются для уже созданных направлений
			transfers = ReDistr.GetTransfersThirdLvl(items, transfers);

			// Выводим отчет для тестов если необходимо
			if (Config.ShowReport)
			{
				Globals.Test.FillListTransfers(transfers, items);
			}

			// Выводим перемещения на лист для перемещений
			Globals.Transfers.FillList(transfers);

			// Выводим параметры отчетов
			Globals.Control.FillReportsParameters();

			// Выбираем лист с перемещениями
			Globals.Transfers.Select();
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

		private void buttonMakeTransfers_Click(object sender, EventArgs e)
		{
			// Архивируем предыдущие перемещения
			ReDistr.ArchiveTransfers();
			
			// Создаем книги для импорта в Excel
			Globals.Transfers.MakeImportTransfers();
		}
	}
}
