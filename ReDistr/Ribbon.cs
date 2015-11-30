using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace ReDistr
{
	public partial class Ribbon
	{
		private void Ribbon_Load(object sender, RibbonUIEventArgs e)
		{
#warning Удалить потом, для отладки
#if(DEBUG)
			// Парсим данные из файлов
			var parser = new Parser();
			Globals.ThisWorkbook.items = parser.Parse(true, true, true);

			var revaluation = ReDistr.GetRevaluations(Globals.ThisWorkbook.items);
			// Заполняем лист с переоценкой
			Globals.Revaluations.FillList(revaluation);

			// Обновляем параметры
			this.UpdateInfo();
#endif
		}

		// Обновить блок с информацией
		public void UpdateInfo()
		{
			labelPeriodSelling.Label = Config.PeriodSellingFrom.ToString("dd.MM.yy") + " - " + Config.PeriodSellingTo.ToString("dd.MM.yy");
			labelPeriodSellingCount.Label = Config.SellingPeriod.ToString() + " (дни)";
			labelStockDate.Label = Config.StockDate.ToString("dd.MM.yy");

			// Включаем/отключаем кнопки в зависимости от результатов парса
			// TODO Доделать с остальными
			if (Config.ParsedStocks & Config.ParsedSealings & Config.ParsedAdditionalParameters)
			{
				buttonGetOrder.Enabled = true;
				buttonGetOrdersLists.Enabled = true;
				buttonGetTransfers.Enabled = true;

			}
			if (Config.ParsedStocks & Config.ParsedSealings & Config.ParsedAdditionalParameters & Config.ParsedCompetitors)
			{
				buttonGetRevaluations.Enabled = true;
			}

		}

#if(!REVAL)
		// Сформировать списки для заказа
		private void buttonGetOrderLists_Click(object sender, RibbonControlEventArgs e)
		{
			var items = Globals.ThisWorkbook.items;

			// Если парсинг не удался, выходим
			if (items == null)
			{
				return;
			}

			// Подготавливаем данные
			ReDistr.PrepareData(items);

			// Выводим таблицу для тестов
			Globals.Test.FillListStocks(items);

			// Обновляем параметры
			this.UpdateInfo();

			// Составляем список для заказа
			var orderRequiredItems = ReDistr.GetOrderRequiredItems(items);

			// Выводим списки для заказа на лист со списками
			Globals.OrderLists.FillList(orderRequiredItems);
		}

		// Сформировать заказы
		private void buttonGetOrders_Click(object sender, RibbonControlEventArgs e)
		{
			var items = Globals.ThisWorkbook.items;

			// Если парсинг не удался, выходим
			if (items == null)
			{
				return;
			}

			// Подготавливаем данные
			ReDistr.PrepareData(items);

			// Выводим таблицу для тестов
			Globals.Test.FillListStocks(items);

			// Обновляем параметры
			this.UpdateInfo();

			// Формирует заказы
			var orders = new List<Order>();
			orders = ReDistr.GetOrders(orders, items);

			// Выводим заказы на страницу заказов
			Globals.Orders.FillList(orders);

			// Выбираем лист с pfrfpfvb
			Globals.Orders.Select();
		}
		// Архивирует старый книги с перемещениями, и создает новые
		private void buttonMakeTransfersBook_Click(object sender, RibbonControlEventArgs e)
		{
			// Архивируем предыдущие перемещения
			ReDistr.ArchiveTransfers();

			// Создаем книги для импорта в Excel
			Globals.Transfers.MakeImportTransfers();
		}

		// Сформировать перемещения
		private void buttonGetTransfers_Click(object sender, RibbonControlEventArgs e)
		{
			var items = Globals.ThisWorkbook.items;

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
			if (Config.StockToTransferSelectedStorageCategory != null)
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

			// Обновляем параметры
			this.UpdateInfo();

			// Выбираем лист с перемещениями
			Globals.Transfers.Select();
		}
#endif
		// Парсим данные
		private void buttonParseData_Click(object sender, RibbonControlEventArgs e)
		{
			// Парсим данные из файлов
			var parser = new Parser();
			Globals.ThisWorkbook.items = parser.Parse(checkBoxIncludeSellings.Checked, checkBoxIncludeAdditionalParameters.Checked, checkBoxIncludeCompetitorsFromAP.Checked);

			// Обновляем параметры
			this.UpdateInfo();
		}

		// Для теста
		private void button1_Click(object sender, RibbonControlEventArgs e)
		{
			button1.Label = Globals.ThisWorkbook.items.Count.ToString();
		}

		// Сформировать переоценку
		private void buttonGetRevaluations_Click(object sender, RibbonControlEventArgs e)
		{
			var revaluation = ReDistr.GetRevaluations(Globals.ThisWorkbook.items);

			// Заполняем лист с переоценкой
			Globals.Revaluations.FillList(revaluation);
		}

	}
}
