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
			buttonGetTransfers_Click(new object(), new EventArgs());
			//buttonMakeTransfersBook_Click(new object(), new EventArgs());
			//buttonGetOrders_Click(new object(), new EventArgs());
			//buttonGetOrderLists_Click(new object(), new EventArgs());
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
			this.buttonGetMoving.Click += new System.EventHandler(this.buttonGetTransfers_Click);
			this.buttonMakeTransfers.Click += new System.EventHandler(this.buttonMakeTransfersBook_Click);
			this.buttonGetOrders.Click += new System.EventHandler(this.buttonGetOrders_Click);
			this.buttonGetOrderLists.Click += new System.EventHandler(this.buttonGetOrderLists_Click);
			this.Startup += new System.EventHandler(this.Лист1_Startup);
			this.Shutdown += new System.EventHandler(this.Лист1_Shutdown);

		}

		#endregion

		private void buttonGetTransfers_Click(object sender, EventArgs e)
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

			// Если необходимо делаем перемещение неликвида на Попова
			if (Config.StockToTransferIlliquid != null)
			{
				transfers = ReDistr.GetTransfersIlliuid(items, transfers);
			}

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

		// Архивирует старый книги с перемещениями, и создает новые
		private void buttonMakeTransfersBook_Click(object sender, EventArgs e)
		{
			// Архивируем предыдущие перемещения
			ReDistr.ArchiveTransfers();

			// Создаем книги для импорта в Excel
			Globals.Transfers.MakeImportTransfers();
		}

		private void buttonGetOrders_Click(object sender, EventArgs e)
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

			// Выводим параметры отчетов
			Globals.Control.FillReportsParameters();

			// Формирует заказы
			var orders = new List<Order>();
			orders = ReDistr.GetOrders(orders, items);

			// Выводим заказы на страницу заказов
			Globals.Orders.FillList(orders);

			// Выбираем лист с pfrfpfvb
			Globals.Orders.Select();
		}

		private void buttonGetOrderLists_Click(object sender, EventArgs e)
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

			// Выводим параметры отчетов
			Globals.Control.FillReportsParameters();

			// Составляем список для заказа
			var orderRequiredItems = ReDistr.GetOrderRequiredItems(items);

			// Выводим списки для заказа на лист со списками
			Globals.OrderLists.FillList(orderRequiredItems);


		}
	}
}
