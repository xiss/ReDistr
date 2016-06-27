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
		// Указываем ячейки с настройками для парсера
		private const string RngNameOfSealingsWb = "B14";
		private const string RngNameOfStocksWb = "B13";
		private const string RngNameListStorageCategoryToTransfers = "B15";
		private const string RngNameOfParamWb = "B16";
		private const string RngNameOfShowReport = "B17";
		private const string RngNameOfMinSoldKits = "B18";
		private const string RngNameOfOnlyPopovaDonor = "B19";
		private const string RngNameFolderTransfer = "B20";
		private const string RngNameFolderArchiveTransfers = "B21";
		private const string RngNameListSelectedStorageCategoryToTransfer = "B22";
		private const string RngNameStockToTransferSelectedStorageCategory = "B23";
		private const string RngNameOfContributorsWb = "B24";
		private const uint RowStartStockCfg = 4; // Строка с которой считываются склады в настройках
		private const string ColStockNameCfg = "A";
		private const string ColStockMinCfg = "B";
		private const string ColStockMaxCfg = "C";
		private const string ColStockSignCfg = "D";
		private const uint RowStartExcludeCompetitors = 4;
		private const string ColExcludeCompetitors = "G";
		private const string RngMinStockForCompetitor = "B25";
		private const string RngWholesaleStock = "B26";
		private const string RngIdPriceAp = "B27";
		private const string RngNameFolderRevaluations = "B28";
		private const string RngNameFolderArchiveRevaluations = "B29";
		// Книга с остатками
		private const uint RowStartStocks = 7; // Строка с которой начинается парсинг остатков
		private const string RngDateStocks = "B3"; // Ячейка содержащая дату отчета
		private const string ColStockSignStocks = "B"; // Колонка содержащая сигнатуру склада
		private const string ColArticleStocks = "C"; // Артикул
		private const string ColCountStocks = "U"; // Остаток
		private const string ColInReserveStocks = "W"; // Резерв
		private const string ColId1CStocks = "X"; // Код товара
		private const string ColStorageCategoryStocks = "Y"; // Категроия хранения
		private const string ColNameStocks = "R"; // Название ЗЧ
		private const string ColManufacturerStocks = "AD"; // Производитель
		private const string ColPriceStocks = "AB"; // Интерет цена
		private const string ColPropertyStocks = "AG"; // Свойство
		private const string ColCostPriceStocks = "V"; // Себестоимость
		// Книга с продажами
		private const uint RowStartSelings = 7; // Строка с которой начинается парсинг продаж
		private const string RngDateSealings = "B3"; // Дата отчета
		private const string ColStockSignSealings = "B"; // Сигнатура склада
		private const string ColArticleSealings = "C"; // Артикул
		private const string ColCountSealings = "V"; // Количество продаж
		private const string ColId1CSealings = "W"; // Код товара
		private const string ColManufacturerSealings = "R"; // Производитель
		private const string ColStorageCategorySealings = "X"; // Категроия хранения
		private const string ColNameSealings = "S"; // Название ЗЧ
		private const string ColPropertySealings = "Y"; // Свойство ЗЧ 2
		// Книга с дополнительными параметрами
		private const uint ListExcludesParameters = 1; // Исключения из перемещений
		private const uint ListInKitParameters = 2; // Кратность
		private const uint ListInBundleParameters = 3; // В упаковке
		private const uint ListSupplierParameters = 4; // Поставщик
		private const uint ListPriceParameters = 5; // Принудительные цены
		private const uint RowStartParameters = 2; // Строка с которой парсится дополнительная информация
		private const string ColId1CParameters = "A"; // Колонка с кодом ЗЧ
		private const string ColArticleParameters = "B";
		private const string ColManufacturerParameters = "D";
		private const string ColNameParameters = "C";
		private const string ColStartParamsParameters = "E"; // Колонка с которой выводится дополнительная информация
		// Книга с конкурентами питерплюса
		private const uint RowStartContributors = 2; // Строка с которой начинается парсинг остатков
		private const string ColArticleContributors = "A"; // Артикул
		private const string ColId1CContributors = "J"; // Код товара
		private const string ColPriceContributors = "C"; // Цена
		private const string ColDeliveryTimeContributors = "E"; // Срок поставки
		private const string ColPositionNumberContributors = "F"; // Строка на портале
		private const string ColRegionContributors = "G"; // Город
		private const string ColIdContributorContributors = "I"; // ID конкурента
		private const string ColCostPriceContributors = "L"; // Себестоимость
		private const string ColCountContributors = "D"; // Остаток у конкурента


		// Диалоговые окна
		private const string MessegeBoxQuestion = "Дата снятия отчета с остатком не соответствует сегодняшней, продолжить?";
		private const string MessegeBoxCaption = "Предупреждение";

		// Получаем параметры с листа настроек
		private void MakeConfig()
		{
			// Выбираем лист с настройками
			Globals.Control.Activate();

			// Настраиваем фабрику
			// Обнуляем параметры
			SimpleStockFactory.CurrentFactory.ClearStockParams();
			var curentRow = RowStartStockCfg;
			uint priority = 1; // Приоритет, от большего к меньшему
			uint count = 0; // Счетчик складов

			while (Globals.Control.Range[ColStockNameCfg + curentRow].Value != null)
			{
				string name = Globals.Control.Range[ColStockNameCfg + curentRow].Value.ToString();
				var minimum = (uint)(Globals.Control.Range[ColStockMinCfg + curentRow].Value);
				var maximum = (uint)(Globals.Control.Range[ColStockMaxCfg + curentRow].Value);
				string signature = Globals.Control.Range[ColStockSignCfg + curentRow].Value.ToString();

				SimpleStockFactory.CurrentFactory.SetStockParams(name, minimum, maximum, signature, priority);
				curentRow++;
				priority++;
				count++;
			}
			Config.StockCount = count;
			Config.SetPossibleTransfers();

			// Конкуренты-исключения
			curentRow = RowStartExcludeCompetitors;
			var list = new List<string>();
			while (Globals.Control.Range[ColExcludeCompetitors + curentRow].Value != null)
			{
				list.Add(Globals.Control.Range[ColExcludeCompetitors + curentRow].Value.ToString());
				curentRow++;
			}
			Config.ListExcludeCompetitors = list;

			// Прописываем в конфиг пути и названия файлов из настроечного листа
			Config.NameOfSealingsWb = Globals.Control.Range[RngNameOfSealingsWb].Value2;
			Config.NameOfStocksWb = Globals.Control.Range[RngNameOfStocksWb].Value2;
			Config.NameOfParametersWb = Globals.Control.Range[RngNameOfParamWb].Value2;
			Config.NameOfCompetitorsWb = Globals.Control.Range[RngNameOfContributorsWb].Value2;
			Config.FolderTransfers = Globals.Control.Range[RngNameFolderTransfer].Value2 + "\\";
			Config.FolderArchiveTransfers = Globals.Control.Range[RngNameFolderArchiveTransfers].Value2 + "\\";
			Config.ShowReport = Globals.Control.Range[RngNameOfShowReport].Value2;
			Config.OneDonor = SimpleStockFactory.CurrentFactory.GetStock(Globals.Control.Range[RngNameOfOnlyPopovaDonor].Value2);
			Config.MinSoldKits = (double)Globals.Control.Range[RngNameOfMinSoldKits].Value2;
			Config.StockToTransferSelectedStorageCategory = SimpleStockFactory.CurrentFactory.GetStock(Globals.Control.Range[RngNameStockToTransferSelectedStorageCategory].Value2);
			Config.WholesaleStock = SimpleStockFactory.CurrentFactory.GetStock(Globals.Control.Range[RngWholesaleStock].Value2);
			Config.MinStockForCompetitor = Globals.Control.Range[RngMinStockForCompetitor].Value2;
			Config.IdPriceAp = Globals.Control.Range[RngIdPriceAp].Value2;
			Config.FolderArchiveRevaluations = Globals.Control.Range[RngNameFolderArchiveRevaluations].Value2;
			Config.FolderRevaluations = Globals.Control.Range[RngNameFolderRevaluations].Value2;
			// Категории для перемещения на указанный склад полностью
			string stringSelectedCategory = Globals.Control.Range[RngNameListSelectedStorageCategoryToTransfer].Value2;
			Config.ListSelectedStorageCategoryToTransfer = stringSelectedCategory.Split(new[] { ';' }).ToList();
			// Категории для перемещения
			string stringCategory = Globals.Control.Range[RngNameListStorageCategoryToTransfers].Value2;
			Config.ListStorageCategoryToTransfers = stringCategory.Split(new[] { ';' }).ToList();
		}

		// Получаем остатки по складам из книги с остатками
		private Dictionary<string, Item> GetItems(out bool _continue)
		{

			_continue = true;
			var items = new Dictionary<string, Item>();

			// Открываем  книгу с остатками
			var fullPath = System.IO.Path.Combine(Globals.ThisWorkbook.Path, "..\\", Config.NameOfStocksWb);
			var stocksWb = Globals.ThisWorkbook.Application.Workbooks.Open(fullPath);

			// Вычисляем дату снятия отчета с остатками
			string dateString = stocksWb.Worksheets[1].Range[RngDateStocks].Value;
			Config.StockDate = DateTime.Parse(dateString.Substring(13, 8));

			// Если дата снятия отчета не равна сегодняшней, предлагаем не продолжать
			if (Config.StockDate != DateTime.Now.Date)
			{
#if(!DEBUG)
				var result = MessageBox.Show(MessegeBoxQuestion, MessegeBoxCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
				if (result == DialogResult.No)
				{

					_continue = false;
					// Закрываем открытую книгу
					stocksWb.Close();
					return null;
				}
#endif
			}

			var curentRow = RowStartStocks;
			var curentStockSignature = String.Empty;

			while (stocksWb.Worksheets[1].Range[ColStockSignStocks + curentRow].Value != null ||
				   stocksWb.Worksheets[1].Range[ColId1CStocks + curentRow].Value != null)
			{
				// Определяем строку с сигнатурой текущего склада
				if (stocksWb.Worksheets[1].Range[ColStockSignStocks + curentRow].Value != null)
				{
					curentStockSignature = stocksWb.Worksheets[1].Range[ColStockSignStocks + curentRow].Value.ToString().ToLower();
					curentRow++;
					continue;
				}

				// Определяем остаток
				double itemCount = 0;
				if (stocksWb.Worksheets[1].Range[ColCountStocks + curentRow].Value is string != true)
				{
					itemCount = stocksWb.Worksheets[1].Range[ColCountStocks + curentRow].Value;
				}

				// Определяем себестоимость
				double itemCostPrice = 0;
				if (stocksWb.Worksheets[1].Range[ColCostPriceStocks + curentRow].Value is string != true & itemCount != 0)
				{
					itemCostPrice = stocksWb.Worksheets[1].Range[ColCostPriceStocks + curentRow].Value / itemCount;
				}

				// Определяем резерв
				double reserveCount = 0;
				if (stocksWb.Worksheets[1].Range[ColInReserveStocks + curentRow].Value is string != true)
				{
					reserveCount = stocksWb.Worksheets[1].Range[ColInReserveStocks + curentRow].Value;
				}

				// Определяем обязательное наличие
				bool requiredAvailability = false;
				if (stocksWb.Worksheets[1].Range[ColStorageCategoryStocks + curentRow].Value is string)
				{
					if (stocksWb.Worksheets[1].Range[ColStorageCategoryStocks + curentRow].Value.ToString() == Config.NameOfStorageCatRequiredAvailability)
					{
						requiredAvailability = true;
					}
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

					// Если СС не установлена, устанавливаем
					//if (item.CostPrice == 0)
					//{
					//	item.CostPrice = itemCostPrice;
					//}
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
						Price = stocksWb.Worksheets[1].Range[ColPriceStocks + curentRow].Value,
						Property = stocksWb.Worksheets[1].Range[ColPropertyStocks + curentRow].Value,
						//CostPrice = itemCostPrice,
						RequiredAvailability = requiredAvailability,
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
					stock.CostPrice = itemCostPrice;
				}
				curentRow++;
			}

			stocksWb.Close();

			return items;
		}

		// Получаем данные по продажам из книги с продажами
		private void GetSellings(Dictionary<string, Item> items)
		{

			// Открываем  книгу с продажами
			var fullPath = System.IO.Path.Combine(Globals.ThisWorkbook.Path, "..\\", Config.NameOfSealingsWb);
			var sellingsWb = Globals.ThisWorkbook.Application.Workbooks.Open(fullPath);

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
					curentStockSignature = sellingsWb.Worksheets[1].Range[ColStockSignSealings + curentRow].Value.ToString().ToLower();
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
					//item.Property = sellingsWb.Worksheets[1].Range[ColPropertySealings + curentRow].Value.ToString();
				}

				// Если не находим создаем ее
				// TODO /1 добавлено, до этого не создавали а переходили к следующей, может настройку сделать?
				if (item == null)
				{
					item = new Item
					{
						Id1C = sellingsWb.Worksheets[1].Range[ColId1CSealings + curentRow].Value.ToString(),
						Article = sellingsWb.Worksheets[1].Range[ColArticleSealings + curentRow].Value.ToString(),
						StorageCategory = sellingsWb.Worksheets[1].Range[ColStorageCategorySealings + curentRow].Value.ToString(),
						Name = sellingsWb.Worksheets[1].Range[ColNameSealings + curentRow].Value.ToString(),
						Manufacturer = sellingsWb.Worksheets[1].Range[ColManufacturerSealings + curentRow].Value.ToString(),
						//Property = sellingsWb.Worksheets[1].Range[ColPropertySealings + curentRow].Value.ToString(),
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
					stock.CountSelings += selingsCount;
				}
				curentRow++;
			}

			sellingsWb.Close();
		}

		// Получаем дополнительные параметры (исключение из перемещений, кратность, в упаковке)
		private void GetAdditionalParameters(Dictionary<string, Item> items)
		{
			// Открываем  книгу с параметрами
			var fullPath = System.IO.Path.Combine(Globals.ThisWorkbook.Path, "..\\", Config.NameOfParametersWb);
			var parametersWb = Globals.ThisWorkbook.Application.Workbooks.Open(fullPath);

			// Обязательное наличие (с созданием карточек)
			// Составляем список складов с листа
			var stockList = new List<string>();
			for (var i = 1; i <= Config.StockCount; i++)
			{
				stockList.Add(parametersWb.Worksheets[ListExcludesParameters].Cells[1, 4 + i].Value.ToLower());
			}

			var curentRow = RowStartParameters;
			// Считываем обязательное наличие
			/*			

						while (parametersWb.Worksheets[ListSupplierParameters].Range[ColId1CParameters + curentRow].Value != null)
						{
							// Ищем запчасть по 1С коду в массиве запчастей
							// Если не находим создаем необходимую ЗЧ
							string curenId1C = parametersWb.Worksheets[ListSupplierParameters].Range[ColId1CParameters + curentRow].Value2;
							if (!items.ContainsKey(curenId1C))
							{
								var item = new Item
								{
									Id1C = parametersWb.Worksheets[ListSupplierParameters].Range[ColId1CParameters + curentRow].Value.ToString(),
									Article = parametersWb.Worksheets[ListSupplierParameters].Range[ColArticleParameters + curentRow].Value,
									//StorageCategory = parametersWb.Worksheets[ListSupplierParameters].Range[ColStorageCategorySealings + curentRow].Value.ToString(),
									Name = parametersWb.Worksheets[ListSupplierParameters].Range[ColNameParameters + curentRow].Value.ToString(),
									Manufacturer = parametersWb.Worksheets[ListSupplierParameters].Range[ColManufacturerParameters + curentRow].Value.ToString()
								};

								// Создаем склады для ЗЧ
								var newStocks = SimpleStockFactory.CurrentFactory.GetAllStocks();
								foreach (var newStock in newStocks)
								{
									item.Stocks.Add(newStock);
								}

								items.Add(item.Id1C, item);
							}
							// Проставляем обязательное наличие у найденной ЗЧ
							foreach (var curentStock in stockList)
							{
								// Проверим, есть ли у текущей ЗЧ текущей склад
								var stock = items[curenId1C].Stocks.Find(s => s.Signature.Contains(curentStock));
								// Если такой склад уже есть, работаем с ним, если нет переходим к следующему складу
								if (stock != null && parametersWb.Worksheets[ListSupplierParameters].Cells[curentRow, 5 + stockList.IndexOf(curentStock)].Value2 == 1)
								{
									stock.RequiredAvailability = true;
								}
							}
							// Проставляем комментарий
							if (items[curenId1C].IsRequiredAvailability())
							{
								items[curenId1C].NoteRequiredAvailability =
									parametersWb.Worksheets[ListSupplierParameters].Cells[curentRow, 5 + Config.StockCount].Value2;
							}

							curentRow++;
						}*/

			// Исключения из перемещений
			// Составляем список складов с листа исключений
			stockList = new List<string>();
			for (var i = 1; i <= Config.StockCount; i++)
			{
				stockList.Add(parametersWb.Worksheets[ListExcludesParameters].Cells[1, 4 + i].Value2.ToLower());
			}

			// Считываем исключения из перемещений
			curentRow = RowStartParameters;

			while (parametersWb.Worksheets[ListExcludesParameters].Range[ColId1CParameters + curentRow].Value2 != null)
			{
				// Ищем запчасть по 1С коду в массиве запчастей
				// Если не находим переходим к следующей строке
				string curenId1C = parametersWb.Worksheets[ListExcludesParameters].Range[ColId1CParameters + curentRow].Value2;
				if (items.ContainsKey(curenId1C))
				{
					// Проставляем исключения у найденной ЗЧ
					foreach (var curentStock in stockList)
					{
						// Проверим, есть ли у текущей ЗЧ текущей склад
						var stock = items[curenId1C].Stocks.Find(s => s.Signature.Contains(curentStock));
						// Если такой склад уже есть, работаем с ним, если нет переходим к следующему складу
						if (stock != null && parametersWb.Worksheets[ListExcludesParameters].Cells[curentRow, 5 + stockList.IndexOf(curentStock)].Value == 1)
						{
							stock.ExcludeFromMoovings = true;
						}
					}
				}

				curentRow++;
			}

			// Кратность запчастей
			curentRow = RowStartParameters;
			while (parametersWb.Worksheets[ListInKitParameters].Range[ColId1CParameters + curentRow].Value2 != null)
			{
				// Ищем запчасть по 1С коду в массиве запчастей
				if (items.ContainsKey(parametersWb.Worksheets[ListInKitParameters].Range[ColId1CParameters + curentRow].Value))
				{
					// Если нашли, проставляем кратность
					Item item = items[parametersWb.Worksheets[ListInKitParameters].Range[ColId1CParameters + curentRow].Value];
					item.InKit = parametersWb.Worksheets[ListInKitParameters].Range[ColStartParamsParameters + curentRow].Value;
				}

				curentRow++;
			}

			// Количество ЗЧ в упаковке
			curentRow = RowStartParameters;
			while (parametersWb.Worksheets[ListInBundleParameters].Range[ColId1CParameters + curentRow].Value2 != null)
			{
				// Ищем запчасть по 1С коду в массиве запчастей
				if (items.ContainsKey(parametersWb.Worksheets[ListInBundleParameters].Range[ColId1CParameters + curentRow].Value))
				{
					// Если нашли, проставляем количество в упаковке
					Item item = items[parametersWb.Worksheets[ListInBundleParameters].Range[ColId1CParameters + curentRow].Value];
					item.InBundle = parametersWb.Worksheets[ListInBundleParameters].Range[ColStartParamsParameters + curentRow].Value;
				}

				curentRow++;
			}

			// Поставщик ЗЧ
			curentRow = RowStartParameters;
			while (parametersWb.Worksheets[ListSupplierParameters].Range[ColId1CParameters + curentRow].Value2 != null)
			{
				// Ищем запчасть по 1С коду в массиве запчастей
				if (items.ContainsKey(parametersWb.Worksheets[ListSupplierParameters].Range[ColId1CParameters + curentRow].Value))
				{
					// Если нашли, проставляем количество в упаковке
					Item item = items[parametersWb.Worksheets[ListSupplierParameters].Range[ColId1CParameters + curentRow].Value];
					item.Supplier = parametersWb.Worksheets[ListSupplierParameters].Range[ColStartParamsParameters + curentRow].Value;
				}

				curentRow++;
			}

			// Цены
			curentRow = RowStartParameters;
			while (parametersWb.Worksheets[ListPriceParameters].Range[ColId1CParameters + curentRow].Value2 != null)
			{
				// Ищем запчасть по 1С коду в массиве запчастей
				if (items.ContainsKey(parametersWb.Worksheets[ListPriceParameters].Range[ColId1CParameters + curentRow].Value))
				{
					// Если нашли, проставляем кратность
					Item item = items[parametersWb.Worksheets[ListPriceParameters].Range[ColId1CParameters + curentRow].Value];
					item.PrePrice = parametersWb.Worksheets[ListPriceParameters].Range[ColStartParamsParameters + curentRow].Value;
				}

				curentRow++;
			}
			parametersWb.Close();
		}

		// Получаем конкурентов на Питерплюсе
		private void GetCompetitorsFromAP(Dictionary<string, Item> items)
		{
			// Открываем  книгу с конкурентами
			var fullPath = System.IO.Path.Combine(Globals.ThisWorkbook.Path, "..\\", Config.NameOfCompetitorsWb);
			var competitorsWb = Globals.ThisWorkbook.Application.Workbooks.Open(fullPath);

			var curentRow = RowStartContributors;

			while (competitorsWb.Worksheets[1].Range[ColId1CContributors + curentRow].Value != null)
			{
				// Если в данных ошибка, переходим к следующей строке
				if (competitorsWb.Worksheets[1].Range[ColId1CContributors + curentRow].Value == "Ошибка")
				{
					curentRow++;
					continue;
				}
				// Ищем запчасть по 1С коду в массиве запчастей, 
				Item item = null;
				if (items.ContainsKey(competitorsWb.Worksheets[1].Range[ColId1CContributors + curentRow].Value.ToString()) == true)
				{
					item = items[competitorsWb.Worksheets[1].Range[ColId1CContributors + curentRow].Value.ToString()];
				}
				// Если не находим то ничего не делаем
				else
				{
					curentRow++;
					continue;
				}
				// Если себестоимость еще не установлена, устанавливаем
				//if (item.CostPrice == 0)
				//{
				//	item.CostPrice = competitorsWb.Worksheets[1].Range[ColCostPriceContributors + curentRow].Value;
				//}
				// Создаем нового конкурента
				var competitor = new Сompetitor
				{
					DeliveryTime = competitorsWb.Worksheets[1].Range[ColDeliveryTimeContributors + curentRow].Value,
					Count = competitorsWb.Worksheets[1].Range[ColCountContributors + curentRow].Value,
					Id = competitorsWb.Worksheets[1].Range[ColIdContributorContributors + curentRow].Value.ToString(),
					PositionNumber = competitorsWb.Worksheets[1].Range[ColPositionNumberContributors + curentRow].Value,
					PriceWithoutAdd = competitorsWb.Worksheets[1].Range[ColPriceContributors + curentRow].Value,
					Region = competitorsWb.Worksheets[1].Range[ColRegionContributors + curentRow].Value,
					Item = item
				};

				item.Сompetitors.Add(competitor);

				curentRow++;
			}

			competitorsWb.Close();

		}

		// Добавляем параметры в конфиг
		public void SetConfig(Dictionary<string, Item> items)
		{
			// Составляем список производителей
			Config.ListSuppliers = items.Values.Select(item => item.Supplier).Distinct().ToList();
		}

		// Основной метод парсера, из него вызываются все остальные
		public Dictionary<string, Item> Parse(bool includeSellings = true, bool includeAdditionalParameters = true, bool includeCompetitorsFromAp = true)
		{
			Globals.ThisWorkbook.items = null;

			// Считываем настройки
			MakeConfig();

			// Создаем список ЗЧ и указываем тукущие остатки
			bool _continue;
			var items = GetItems(out _continue);
			Config.ParsedStocks = true;

			// Если отчет не подходит, выходим
			if (!_continue) return null;

			// Добавляем информацию по продажам
			if (includeSellings)
			{
				GetSellings(items);
				Config.ParsedSealings = true;
			}

			// Добавляем Кратность и исключения
			if (includeAdditionalParameters)
			{
				GetAdditionalParameters(items);
				Config.ParsedAdditionalParameters = true;
			}

			// Добавляем конкурентов
			if (includeCompetitorsFromAp)
			{
				GetCompetitorsFromAP(items);
				Config.ParsedCompetitors = true;
			}

			// Настраиваем конфиг
			SetConfig(items);

			return items;
		}
	}
}
