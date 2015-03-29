﻿using System;
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
		// TODO Убрать контрол, он похоже не нужен
		private readonly Control _control;

		// Конструткор
		public Parser(Control control)
		{
			_control = control;
		}

		// Указываем ячейки с настройками для парсера
		private const string RngNameOfSealingsWb = "B14";
		private const string RngNameOfStocksWb = "B13";
		private const string RngPuthToThisWb = "B15";
		private const string RngNameOfParamWb = "B16";
		private const string RngNameOfShowReport = "B17";
		private const string RngNameOfMinSoldKits = "B18";
		private const string RngNameOfOnlyPopovaDonor = "B19";
		private const string RngNameFolderTransfer = "B20";
		private const uint RowStartStockCfg = 4; // Строка с которой считываются склады в настройках
		private const string ColStockNameCfg = "A";
		private const string ColStockMinCfg = "B";
		private const string ColStockMaxCfg = "C";
		private const string ColStockSignCfg = "D";
		// Книга с остатками
		private const uint RowStartStocks = 7; // Строка с которой начинается парсинг остатков
		private const string RngDateStocks = "B3"; // Ячейка содержащая дату отчета
		private const string ColStockSignStocks = "B"; // Колонка содержащая сигнатуру склада
		private const string ColArticleStocks = "C"; // Артикул
		private const string ColCountStocks = "U"; // Остаток
		private const string ColInReserveStocks = "Z"; // Резерв
		private const string ColId1CStocks = "W"; // Код товара
		private const string ColStorageCategoryStocks = "X"; // Категроия хранения
		private const string ColNameStocks = "R"; // Название ЗЧ
		private const string ColManufacturerStocks = "AB"; // Производитель
		// Книга с продажами
		private const uint RowStartSelings = 7; // Строка с которой начинается парсинг продаж
		private const string RngDateSealings = "B3"; // Дата отчета
		private const string ColStockSignSealings = "B"; // Сигнатура склада
		private const string ColArticleSealings = "C"; // Артикул
		private const string ColCountSealings = "Y"; // Количество продаж
		private const string ColId1CSealings = "S"; // Код товара
		// Книга с дополнительными параметрами
		private const uint RowStartParameters = 2; // Строка с которой парсится дополнительная информация
		private const string ColId1CParameters = "A"; // Колонка с кодом ЗЧ
		private const string ColStartParamsParameters = "E"; // Колонка с которой выводится дополнительная информация
		// Диалоговые окна
		private const string MessegeBoxQuestion = "Дата снятия отчета с остатком не соответствует сегодняшней, продолжить?";
		private const string MessegeBoxCaption = "Предупреждение";

		// Получаем параметры с листа настроек
		private void MakeConfig()
		{
			// Выбираем лист с настройками
			Globals.Control.Activate();

			// Прописываем в конфиг пути и названия файло виз настроечного листа
			Config.NameOfSealingsWb = _control.Range[RngNameOfSealingsWb].Value2;
			Config.NameOfStocksWb = _control.Range[RngNameOfStocksWb].Value2;
			Config.PuthToThisWb = _control.Range[RngPuthToThisWb].Value2;
			Config.NameOfParametersWb = _control.Range[RngNameOfParamWb].Value2;
			Config.FolderTransfers = _control.Range[RngNameFolderTransfer].Value2 + "\\";
			Config.ShowReport = _control.Range[RngNameOfShowReport].Value2;
			Config.OnlyPopovaDonor = _control.Range[RngNameOfOnlyPopovaDonor].Value2;
			Config.MinSoldKits = (double)_control.Range[RngNameOfMinSoldKits].Value2;

			// Настраиваем фабрику
			var curentRow = RowStartStockCfg;
			uint priority = 1; // Приоритет, от большего к меньшему
			uint count = 0; // Счетчик складов

			while (_control.Range[ColStockNameCfg + curentRow].Value != null)
			{
				string name = _control.Range[ColStockNameCfg + curentRow].Value.ToString();
				var minimum = (uint)(_control.Range[ColStockMinCfg + curentRow].Value);
				var maximum = (uint)(_control.Range[ColStockMaxCfg + curentRow].Value);
				string signature = _control.Range[ColStockSignCfg + curentRow].Value.ToString();

				SimpleStockFactory.CurrentFactory.SetStockParams(name, minimum, maximum, signature, priority);
				curentRow++;
				priority++;
				count++;
			}
			Config.StockCount = count;
			Config.SetPossibleTransfers();
			
		}

		// Получаем остатки по складам
		private Dictionary<string, Item> GetItems(out bool _continue)
		{

			_continue = true;
			var items = new Dictionary<string, Item>();

			// Открываем  книгу с остатками
			var stocksWb = _control.Application.Workbooks.Open(Config.PuthToThisWb + Config.NameOfStocksWb);

			// Вычисляем дату снятия отчета с остатками
			string dateString = stocksWb.Worksheets[1].Range[RngDateStocks].Value;
			Config.StockDate = DateTime.Parse(dateString.Substring(13, 8));

			// Если дата снятия отчета не равна сегодняшней, предлагаем не продолжать
			if (Config.StockDate != new DateTime().Date)
			{
#if(!DEBUG)
				var result = MessageBox.Show(MessegeBoxQuestion, MessegeBoxCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
				if (result == DialogResult.No)
				{

					_continue = false;
					return null;
				}
#endif
			}

			var curentRow = RowStartStocks;
			var curentStockSignature = String.Empty;

			while (stocksWb.Worksheets[1].Range[ColStockSignStocks + curentRow].Value != null || stocksWb.Worksheets[1].Range[ColId1CStocks + curentRow].Value != null)
			{
				// Определяем строку с сигнатурой текущего склада
				if (stocksWb.Worksheets[1].Range[ColStockSignStocks + curentRow].Value != null)
				{
					curentStockSignature = stocksWb.Worksheets[1].Range[ColStockSignStocks + curentRow].Value.ToString();
					curentRow++;
					continue;
				}

				// Определяем остаток
				double itemCount = 0;
				if (stocksWb.Worksheets[1].Range[ColCountStocks + curentRow].Value is string != true)
				{
					itemCount = stocksWb.Worksheets[1].Range[ColCountStocks + curentRow].Value;
				}

				// Определяем резерв
				double reserveCount = 0;
				if (stocksWb.Worksheets[1].Range[ColInReserveStocks + curentRow].Value is string != true)
				{
					reserveCount = stocksWb.Worksheets[1].Range[ColInReserveStocks + curentRow].Value;
				}

				// Если резерв отрицателен, округляем его до 0
				if (reserveCount < 0)
				{
					reserveCount = 0;
				}

				// Ищем запчасть по 1С коду в массиве запчастей
				Item item = null;
				if (items.ContainsKey(stocksWb.Worksheets[1].Range[ColId1CStocks + curentRow].Value.ToString()))
				{
					item = items[stocksWb.Worksheets[1].Range[ColId1CStocks + curentRow].Value.ToString()];
				}

				// Если не находим, создаем ее
				if (item == null)
				{
					item = new Item
					{
						Id1C = stocksWb.Worksheets[1].Range[ColId1CStocks + curentRow].Value.ToString(),
						Article = stocksWb.Worksheets[1].Range[ColArticleStocks + curentRow].Value.ToString(),
						StorageCategory = stocksWb.Worksheets[1].Range[ColStorageCategoryStocks + curentRow].Value.ToString(),
						Name = stocksWb.Worksheets[1].Range[ColNameStocks + curentRow].Value.ToString(),
						Manufacturer = stocksWb.Worksheets[1].Range[ColManufacturerStocks + curentRow].Value.ToString(),
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
					stock.SetOriginCount(itemCount);
					stock.InReserve = reserveCount;
				}
				curentRow++;
			}

			stocksWb.Close();

			return items;
		}

		// Получаем данные по продажам
		private void GetSellings(Dictionary<string, Item> items)
		{

			// Открываем  книгу с продажами
			var sellingsWb = _control.Application.Workbooks.Open(Config.PuthToThisWb + Config.NameOfSealingsWb);

			// Вычисляем начальную и конечную дату периода продаж
			string dateString = sellingsWb.Worksheets[1].Range[RngDateSealings].Value;
			Config.PeriodSellingFrom = DateTime.Parse(dateString.Substring(12, 8));
			Config.PeriodSellingTo = DateTime.Parse(dateString.Substring(23, 8));
			Config.SellingPeriod = (Config.PeriodSellingTo - Config.PeriodSellingFrom).Days;

			var curentRow = RowStartSelings;
			var curentStockSignature = String.Empty;

			while (sellingsWb.Worksheets[1].Range[ColStockSignSealings + curentRow].Value != null || sellingsWb.Worksheets[1].Range[ColId1CSealings + curentRow].Value != null)
			{
				// Определяем строку с сигнатурой текущего склада
				if (sellingsWb.Worksheets[1].Range[ColStockSignSealings + curentRow].Value != null)
				{
					curentStockSignature = sellingsWb.Worksheets[1].Range[ColStockSignSealings + curentRow].Value.ToString();
					curentRow++;
					continue;
				}

				// Определяем продажи
				double selingsCount = 0;
				if (sellingsWb.Worksheets[1].Range[ColCountSealings + curentRow].Value is string != true)
				{
					selingsCount = sellingsWb.Worksheets[1].Range[ColCountSealings + curentRow].Value;
				}

				// Ищем запчасть по 1С коду в массиве запчастей
				Item item = null;
				if (items.ContainsKey(sellingsWb.Worksheets[1].Range[ColId1CSealings + curentRow].Value.ToString()) == true)
				{
					item = items[sellingsWb.Worksheets[1].Range[ColId1CSealings + curentRow].Value.ToString()];
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
			var parametersWb = _control.Application.Workbooks.Open(Config.PuthToThisWb + Config.NameOfParametersWb);

			// Исключения из перемещений
			// Составляем список складов с листа исключений
			var stockList = new List<string>();
			for (var i = 1; i <= Config.StockCount; i++)
			{
				stockList.Add(parametersWb.Worksheets[1].Cells[1, 4 + i].Value);
			}

			// Считываем исключения из перемещений
			var curentRow = RowStartParameters;
			while (parametersWb.Worksheets[1].Range[ColId1CParameters + curentRow].Value != null)
			{
				// Ищем запчасть по 1С коду в массиве запчастей
				// Если не находим переходим к следующей строке
				string curenId1C = parametersWb.Worksheets[1].Range[ColId1CParameters + curentRow].Value;
				if (items.ContainsKey(curenId1C))
				{
					// Проставляем исключения у найденной ЗЧ
					foreach (var curentStock in stockList)
					{
						// Проверим, есть ли у текущей ЗЧ текущей склад
						var stock = items[curenId1C].Stocks.Find(s => s.Signature.Contains(curentStock));
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
			curentRow = RowStartParameters;
			while (parametersWb.Worksheets[2].Range[ColId1CParameters + curentRow].Value != null)
			{
				// Ищем запчасть по 1С коду в массиве запчастей
				if (items.ContainsKey(parametersWb.Worksheets[2].Range[ColId1CParameters + curentRow].Value))
				{
					// Если нашли, проставляем кратность
					Item item = items[parametersWb.Worksheets[2].Range[ColId1CParameters + curentRow].Value];
					item.InKit = parametersWb.Worksheets[2].Range[ColStartParamsParameters + curentRow].Value;
				}

				curentRow++;
			}

			// Количество ЗЧ в упаковке
			curentRow = RowStartParameters;
			while (parametersWb.Worksheets[3].Range[ColId1CParameters + curentRow].Value != null)
			{
				// Ищем запчасть по 1С коду в массиве запчастей
				if (items.ContainsKey(parametersWb.Worksheets[3].Range[ColId1CParameters + curentRow].Value))
				{
					// Если нашли, проставляем количество в упаковке
					Item item = items[parametersWb.Worksheets[3].Range[ColId1CParameters + curentRow].Value];
					item.InBundle = parametersWb.Worksheets[3].Range[ColStartParamsParameters + curentRow].Value;
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
