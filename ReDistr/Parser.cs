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
	class Parser
	{
		private readonly Control _control;

		// Конструткор
		public Parser(Control control)
		{
			_control = control;
		}

		// Указываем ячейки с настройками для парсера
		private const string RangeNameOfSealingsWb = "B14";
		private const string RangeNameOfStocksWb = "B13";
		private const string RangePuthToThisWb = "B15";
		private const string RangeNameOfParametersWb = "B16";
		private const uint RowNumberStockConfig = 4;
		private const uint RowNumberStocks = 7;
		private const uint RowNumberSelings = 7;
		private const uint RowNumberParameters = 2;
		private const string MessegeBoxQuestion = "Дата снятия отчета с остатком не соответствует сегодняшней, продолжить?";
		private const string MessegeBoxCaption = "Предупреждение";

		// Получаем параметры с листа настроек
		private void MakeConfig()
		{
			// Выбираем лист с настройками
			Globals.Control.Activate();

			// Прописываем в конфик пути и названия файло виз настроечного листа
			Config.NameOfSealingsWB = _control.Range[RangeNameOfSealingsWb].Value2;
			Config.NameOfStocksWB = _control.Range[RangeNameOfStocksWb].Value2;
			Config.PuthToThisWB = _control.Range[RangePuthToThisWb].Value2;
			Config.NameOfParametersWb = _control.Range[RangeNameOfParametersWb].Value2;

			// Настраиваем фабрику
			var curentRow = RowNumberStockConfig;
			uint priority = 1; // Приоритет, от большего к меньшему
			uint count = 0; // Счетчик складов

			while (_control.Range["A" + curentRow].Value != null)
			{
				string name = _control.Range["A" + curentRow].Value.ToString();
				var minimum = (uint)(_control.Range["B" + curentRow].Value);
				var maximum = (uint)(_control.Range["C" + curentRow].Value);
				string signature = _control.Range["D" + curentRow].Value.ToString();

				SimpleStockFactory.CurrentFactory.SetStockParams(name, minimum, maximum, signature, priority);
				curentRow++;
				priority++;
				count++;
			}
			Config.StockCount = count;
		}

		// Получаем остатки по складам
		private Dictionary<string, Item> GetItems(out bool _continue)
		{

			_continue = true;
			var items = new Dictionary<string, Item>();

			// Открываем  книгу с остатками
			var stocksWb = _control.Application.Workbooks.Open(Config.PuthToThisWB + Config.NameOfStocksWB);

			// Вычисляем дату снятия отчета с остатками
			string dateString = stocksWb.Worksheets[1].Range["B3"].Value;
			Config.StockDate = DateTime.Parse(dateString.Substring(13, 8));

			// Если дата снятия отчета не равна сегодняшней, предлагаем не продолжать
			if (Config.StockDate != new DateTime().Date)
			{
				var result = MessageBox.Show(MessegeBoxQuestion, MessegeBoxCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
				if (result == DialogResult.No)
				{
					_continue = false;
					return null;
				}
			}

			var curentRow = RowNumberStocks;
			var curentStockSignature = String.Empty;

			while (stocksWb.Worksheets[1].Range["B" + curentRow].Value != null || stocksWb.Worksheets[1].Range["C" + curentRow].Value != null)
			{
				// Определяем строку с сигнатурой текущего склада
				if (stocksWb.Worksheets[1].Application.Range["B" + curentRow].Value != null)
				{
					curentStockSignature = stocksWb.Worksheets[1].Application.Range["B" + curentRow].Value.ToString();
					curentRow++;
					continue;
				}

				// Определяем остаток
				double itemCount = 0;
				if (stocksWb.Worksheets[1].Application.Range["U" + curentRow].Value is string != true)
				{
					itemCount = stocksWb.Worksheets[1].Application.Range["U" + curentRow].Value;
				}

				// Определяем резерв
				double reserveCount = 0;
				if (stocksWb.Worksheets[1].Application.Range["Y" + curentRow].Value is string != true)
				{
					reserveCount = stocksWb.Worksheets[1].Application.Range["Y" + curentRow].Value;
				}

				// Если резерв отрицателен, округляем его до 0
				if (reserveCount < 0)
				{
					reserveCount = 0;
				}

				// Ищем запчасть по 1С коду в массиве запчастей
				Item item = null;
				if (items.ContainsKey(stocksWb.Worksheets[1].Application.Range["V" + curentRow].Value) == true)
				{
					item = items[stocksWb.Worksheets[1].Application.Range["V" + curentRow].Value];
				}

				// Если не находим, создаем ее
				if (item == null)
				{
					item = new Item
					{
						Id1C = stocksWb.Worksheets[1].Application.Range["V" + curentRow].Value,
						Article = stocksWb.Worksheets[1].Application.Range["C" + curentRow].Value,
						StorageCategory = stocksWb.Worksheets[1].Application.Range["W" + curentRow].Value,
						Name = stocksWb.Worksheets[1].Application.Range["R" + curentRow].Value,
						Manufacturer = stocksWb.Worksheets[1].Application.Range["AA" + curentRow].Value,
					};

					// Создаем склады для ЗЧ
					var newStocks = SimpleStockFactory.CurrentFactory.GetAllStocks();
					foreach (var newStock in newStocks)
					{
						item.Stocks.Add(newStock);
					}

					items.Add(item.Id1C, item);
				}

				// Проверим, есть ли у текущей ЗЧ текущей склад
				var stock = item.Stocks.Find(s => curentStockSignature.Contains(s.Signature));

				// Если такой склад уже есть, работаем с ним
				if (stock != null)
				{
					stock.Count += itemCount;
					stock.InReserve += reserveCount;
				}

				// Если склада нет, создаем его
				//else
				//{
				//	stock = SimpleStockFactory.CurrentFactory.TryGetStock(curentStockSignature);
				//	stock.Count = itemCount;
				//	stock.InReserve = reserveCount;
				//	item.Stocks.Add(stock);
				//}

				curentRow++;
			}

			stocksWb.Close();
			return items;

		}

		// Получаем данные по продажам
		private void GetSellings(Dictionary<string, Item> items)
		{

			// Открываем  книгу с продажами
			var sellingsWb = _control.Application.Workbooks.Open(Config.PuthToThisWB + Config.NameOfSealingsWB);

			// Вычисляем начальную и конечную дату периода продаж
			string dateString = sellingsWb.Worksheets[1].Range["B3"].Value;
			Config.periodSellingFrom = DateTime.Parse(dateString.Substring(12, 8));
			Config.periodSellingTo = DateTime.Parse(dateString.Substring(23, 8));
			Config.sellingPeriod = (Config.periodSellingTo - Config.periodSellingFrom).Days;

			var curentRow = RowNumberSelings;
			var curentStockSignature = String.Empty;

			while (sellingsWb.Worksheets[1].Range["B" + curentRow].Value != null || sellingsWb.Worksheets[1].Range["C" + curentRow].Value != null)
			{
				// Определяем строку с сигнатурой текущего склада
				if (sellingsWb.Worksheets[1].Application.Range["B" + curentRow].Value != null)
				{
					curentStockSignature = sellingsWb.Worksheets[1].Application.Range["B" + curentRow].Value.ToString();
					curentRow++;
					continue;
				}

				// Определяем продажи
				double selingsCount = 0;
				if (sellingsWb.Worksheets[1].Application.Range["X" + curentRow].Value is string != true)
				{
					selingsCount = sellingsWb.Worksheets[1].Application.Range["X" + curentRow].Value;
				}

				// Ищем запчасть по 1С коду в массиве запчастей
				Item item = null;
				if (items.ContainsKey(sellingsWb.Worksheets[1].Application.Range["S" + curentRow].Value) == true)
				{
					item = items[sellingsWb.Worksheets[1].Application.Range["S" + curentRow].Value];
				}

				// Если не находим переходим к следующей строке
				if (item == null)
				{
					curentRow++;
					continue;
				}

				// Проверим, есть ли у текущей ЗЧ текущей склад
				var stock = item.Stocks.Find(s => curentStockSignature.Contains(s.Signature));

				// Если такой склад уже есть, работаем с ним
				if (stock != null)
				{
					stock.SelingsCount += selingsCount;
				}

				// Если склада нет, создаем его
				else
				{
					stock = SimpleStockFactory.CurrentFactory.TryGetStock(curentStockSignature);
					stock.SelingsCount = selingsCount;
					item.Stocks.Add(stock);
				}
				curentRow++;
			}

			sellingsWb.Close();
		}

		// Получаем дополнительные параметры (исключение из перемещений, кратность, в упаковке)
		private void GetAdditionalParameters(Dictionary<string, Item> items)
		{
			// Открываем  книгу с параметрами
			var parametersWb = _control.Application.Workbooks.Open(Config.PuthToThisWB + Config.NameOfParametersWb);

			// Исключения из перемещений
			//Определяем список известных складов
			var stockList = new List<string>();
			for (var i = 1; i <= Config.StockCount; i++)
			{
				stockList.Add(parametersWb.Worksheets[1].Cells[1, 4 + i].Value);
			}

			// Считываем исключения из перемещений
			var curentRow = RowNumberParameters;
			while (parametersWb.Worksheets[1].Range["A" + curentRow].Value != null)
			{
				// Ищем запчасть по 1С коду в массиве запчастей
				// Если не находим переходим к следующей строке
				if (items.ContainsKey(parametersWb.Worksheets[1].Application.Range["A" + curentRow].Value))
				{
					Item item = items[parametersWb.Worksheets[1].Application.Range["A" + curentRow].Value];

					// Проставляем исключения у найденной ЗЧ
					foreach (var curentStock in stockList)
					{
						// Проверим, есть ли у текущей ЗЧ текущей склад
						var stock = item.Stocks.Find(s => s.Signature.Contains(curentStock));
						// Если такой склад уже есть, работаем с ним, если нет переходим к следующему складу
						if (stock != null && parametersWb.Worksheets[1].Cells[curentRow, 5 + stockList.IndexOf(curentStock)].Value == 1)
						{
							stock.ExcludeFromMoovings = true;
						}
					}
				}

				curentRow++;
			}

			// Кратность запчастей
			curentRow = RowNumberParameters;
			while (parametersWb.Worksheets[2].Range["A" + curentRow].Value != null)
			{
				// Ищем запчасть по 1С коду в массиве запчастей
				if (items.ContainsKey(parametersWb.Worksheets[2].Range["A" + curentRow].Value))
				{
					// Если нашли, проставляем кратность
					Item item = items[parametersWb.Worksheets[2].Range["A" + curentRow].Value];
					item.inKit = parametersWb.Worksheets[2].Range["E" + curentRow].Value;
				}

				curentRow++;
			}

			// Количество ЗЧ в упаковке
			curentRow = RowNumberParameters;
			while (parametersWb.Worksheets[3].Range["A" + curentRow].Value != null)
			{
				// Ищем запчасть по 1С коду в массиве запчастей
				if (items.ContainsKey(parametersWb.Worksheets[3].Range["A" + curentRow].Value))
				{
					// Если нашли, проставляем количество в упаковке
					Item item = items[parametersWb.Worksheets[3].Range["A" + curentRow].Value];
					item.inBundle = parametersWb.Worksheets[3].Range["E" + curentRow].Value;
				}

				curentRow++;
			}
			parametersWb.Close();
		}

		// Основной метод парсера, из него вызываются все остальные
		public Dictionary<string, Item> Parse()
		{
			// Считываем настройки
			MakeConfig();

			// Создаем список ЗЧ и указываем тукущие остатки
			bool _continue;
			var items = GetItems(out _continue);
			// Если отчет не подходит, выходим
			if (!_continue) return null;

			// Добавляем информацию по продажам
			GetSellings(items);

			// Добавляем Кратность и исключения
			GetAdditionalParameters(items);

			return items;
		}
	}
}
