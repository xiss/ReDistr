using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace ReDistr
{
	class ReDistr
	{
		// Расчитывает первичные параметры для перемещений
		public static void PrepareData(Dictionary<string, Item> items)
		{
			foreach (var item in items)
			{
				foreach (var stock in item.Value.Stocks)
				{
					stock.UpdateSailPersent(item.Value);
					stock.UpdateMinStock(item.Value);
					stock.UpdateMaxStock(item.Value);
				}
			}
		}

		// Создает необходимые перемещения для обеспечения одного комплекта на складах с ненулевым миностатком
		public static List<Transfer> GetTransfersFirstLvl(Dictionary<string, Item> items, List<Transfer> transfers)
		{
			// Перебираем список ЗЧ
			foreach (KeyValuePair<string, Item> item in items)
			{
				// Сортируем склады по приоритету, в первую очередб обрабатываем более приоритетные
				item.Value.Stocks = item.Value.Stocks.OrderByDescending(stock => stock.Priority).ToList();
				// Расчитываем свободные остатки на складах
				item.Value.UpdateFreeStocks("kit");

				// Перебираем список складов у ЗЧ
				foreach (var stock in item.Value.Stocks)
				{
					// Определяем, требуется ли перемещения для данной категории хранения и установлен ли у ЗЧ параметр RequiredAvailability, 
					// в последнем случае перемещение все равно делаем
					if (!Config.ListStorageCategoryToTransfers.Contains(item.Value.StorageCategory) && !item.Value.RequiredAvailability)
					{
						continue;
					}
					// Если данный склад исключен из перемещений, переходим к следующему складу
					if (stock.ExcludeFromMoovings)
					{
						continue;
					}
					// Определяем количество для обеспечения одного комплекта, если потребности нет, переходим к следующему складу
					var need = stock.GetNeedToInKit(item.Value);
					var possibleDonors = item.Value.GetListOfPossibleDonors();
					if (need == 0)
					{
						continue;
					}
					// Определяем есть ли доноры, если доноров нет, переходим к следующей ЗЧ
					if (possibleDonors == null)
					{
						break;
					}
					// Определяем, достаточно ли общего свободного остатка для обеспечения потребности, если нет, переходим к следующему складу
					if (need > item.Value.GetSumFreeStocks())
					{
						continue;
					}
					// Создаем необходимые перемещения от доноров
					foreach (var possibleDonor in possibleDonors)
					{
						// Если потребность удовлетворяется одним донором
						if (need <= possibleDonor.FreeStock)
						{
							var transfer = new Transfer
							{
								StockFrom = possibleDonor,
								StockTo = stock,
								Count = need,
								Item = item.Value
							};
							transfer.Apply();
							transfers.Add(transfer);

							break;
						}
						// Если нужны еще доноры
						else
						{
							var transfer = new Transfer
							{
								StockFrom = possibleDonor,
								StockTo = stock,
								Count = possibleDonor.FreeStock,
								Item = item.Value
							};
							need -= possibleDonor.FreeStock;
							transfer.Apply();
							transfers.Add(transfer);
						}
					}
				}
			}
			return transfers;
		}

		// Создает необходимые пеермещения, для обеспечения минимального остатка
		public static List<Transfer> GetTransfersSecondLvl(Dictionary<string, Item> items, List<Transfer> transfers)
		{
			// Перебираем список ЗЧ
			foreach (KeyValuePair<string, Item> item in items)
			{
				// Сортируем склады по приоритету, в первую очередб обрабатываем более приоритетные
				item.Value.Stocks = item.Value.Stocks.OrderByDescending(stock => stock.Priority).ToList();
				// Расчитываем свободные остатки на складах
				item.Value.UpdateFreeStocks("minStock");

				// Перебираем список складов у ЗЧ
				foreach (var stock in item.Value.Stocks)
				{
					// Определяем, требуется ли перемещения для данной категории хранения
					if (!Config.ListStorageCategoryToTransfers.Contains(item.Value.StorageCategory))
					{
						continue;
					}

					// Если данный склад исключен из перемещений, переходим к следующему складу
					if (stock.ExcludeFromMoovings)
					{
						continue;
					}

					// Определяем количество для обеспечения мин остатка, если потребности нет, переходим к следующему складу
					var need = stock.GetNeedToMinStock(item.Value);
					var possibleDonors = item.Value.GetListOfPossibleDonors();
					// Определяем, достаточно ли общего свободного остатка для обеспечения потребности, если нет, уменьшаем потребность ориентируясь на кратность
					if (need > item.Value.GetSumFreeStocks())
					{
						need = Math.Floor((stock.Count + item.Value.GetSumFreeStocks()) / item.Value.InKit) * item.Value.InKit - stock.Count;
					}
					// Если потребность нулевая или отрицательная, переходим к следующему складу
					if (need <= 0)
					{
						continue;
					}
					// Определяем есть ли доноры, если доноров нет, переходим к следующей ЗЧ
					if (possibleDonors == null)
					{
						break;
					}

					// Создаем необходимые перемещения от доноров
					foreach (var possibleDonor in possibleDonors)
					{
						// Если потребность удовлетворяется одним донором
						if (need <= possibleDonor.FreeStock)
						{
							var transfer = new Transfer
							{
								StockFrom = possibleDonor,
								StockTo = stock,
								Count = need,
								Item = item.Value
							};
							transfer.Apply();
							transfers.Add(transfer);

							break;
						}
						// Если нужны еще доноры
						else
						{
							var transfer = new Transfer
							{
								StockFrom = possibleDonor,
								StockTo = stock,
								Count = possibleDonor.FreeStock,
								//Count = need,
								Item = item.Value
							};
							need -= possibleDonor.FreeStock;
							transfer.Apply();
							transfers.Add(transfer);
						}
					}
				}
			}
			return transfers;
		}

		// TODO следующей итерацией мы обеспечиваем наличия мин остатка, но не обеспечиваем перемещение в случае если мин остаток равен остатку. Пока так и оставлю.
		// Создает необходимые перемещения для обеспечения необходимого запаса
		public static List<Transfer> GetTransfersThirdLvl(Dictionary<string, Item> items, List<Transfer> transfers)
		{
			// Перебираем список ЗЧ
			foreach (KeyValuePair<string, Item> item in items)
			{
				// Сортируем склады по приоритету, в первую очередь обрабатываем более приоритетные
				item.Value.Stocks = item.Value.Stocks.OrderByDescending(stock => stock.Priority).ToList();
				// Расчитываем свободные остатки на складах
				item.Value.UpdateFreeStocks("minStock");

				// Перебираем список складов у ЗЧ
				foreach (var stock in item.Value.Stocks)
				{
					// Определяем, требуется ли перемещения для данной категории хранения
					if (!Config.ListStorageCategoryToTransfers.Contains(item.Value.StorageCategory))
					{
						continue;
					}

					// Если данный склад исключен из перемещений, переходим к следующему складу
					if (stock.ExcludeFromMoovings)
					{
						continue;
					}

					// Определяем количество для перемещения, если перемещать ничего не нужно переходим к следующему складу
					var need = stock.GetNeedToSafety(item.Value);
					// TODO /9 Нужно оставить только уникальные перемещения по донору (Проверить)
					var existTransfers = transfers.Where(transfer => transfer.Item == item.Value && transfer.StockTo == stock).Distinct().ToList();

					var possibleDonors = item.Value.GetListOfPossibleDonors(existTransfers);
					if (need == 0)
					{
						continue;
					}
					// Определяем есть ли доноры, если доноров нет, переходим к следующей ЗЧ
					if (possibleDonors == null)
					{
						break;
					}
					// Определяем, достаточно ли общего свободного остатка для обеспечения потребности, если нет, переходим к следующему складу
					if (need > item.Value.GetSumFreeStocks(existTransfers))
					{
						need = Math.Floor((stock.Count + item.Value.GetSumFreeStocks(existTransfers)) / item.Value.InKit) * item.Value.InKit - stock.Count;
					}
					// Создаем необходимые перемещения от доноров
					foreach (var possibleDonor in possibleDonors)
					{
						// Если потребность удовлетворяется одним донором
						if (need <= possibleDonor.FreeStock)
						{
							var transfer = new Transfer
							{
								StockFrom = possibleDonor,
								StockTo = stock,
								Count = need,
								Item = item.Value
							};
							transfer.Apply();
							transfers.Add(transfer);

							break;
						}
						// Если нужны еще доноры
						else
						{
							var transfer = new Transfer
							{
								StockFrom = possibleDonor,
								StockTo = stock,
								Count = possibleDonor.FreeStock,
								Item = item.Value
							};
							need -= possibleDonor.FreeStock;
							transfer.Apply();
							transfers.Add(transfer);
						}
					}
				}

			}
			return transfers;
		}

		// Создает перемещения неликвида на попова если там 0
		public static List<Transfer> GetTransfersIlliuid(Dictionary<string, Item> items, List<Transfer> transfers)
		{
			var stockTo = Config.StockToTransferSelectedStorageCategory;

			// Перебираем список ЗЧ
			foreach (KeyValuePair<string, Item> item in items)
			{
				// Расчитываем свободные остатки на складах
				// TODO наверное убрать, не помню
				//item.Value.UpdateFreeStocks("kit");

				// Если категория не указана в списке, переходим к следующей ЗЧ
				if (!Config.ListSelectedStorageCategoryToTransfer.Contains(item.Value.StorageCategory))
				{
					continue;
				}

				// Если на нужном складе 0, а на других есть в наличии, делаем перемещение
				if (!(item.Value.Stocks.Find(stock => stock == stockTo).Count == 0 && item.Value.GetSumStocks() > 0))
				{
					continue;
				}

				// Перебираем список складов у ЗЧ
				foreach (var stock in item.Value.Stocks)
				{
					// Если данный склад исключен из перемещений, переходим к следующему складу
					if (stock.ExcludeFromMoovings)
					{
						continue;
					}

					// Если на данном складе установлено обязательное наличие, Переходим к следующему
					if (item.Value.RequiredAvailability)
					{
						continue;
					}

					var possibleDonors = item.Value.GetListOfPossibleDonors();
					// Создаем необходимые перемещения от доноров
					foreach (var possibleDonor in possibleDonors)
					{
						var transfer = new Transfer
							{
								StockFrom = possibleDonor,
								StockTo = stockTo,
								Count = possibleDonor.FreeStock,
								Item = item.Value
							};
						transfer.Apply();
						transfers.Add(transfer);
					}
				}
			}
			return transfers;
		}

		// Дает список возможных перемещений
		public static List<Transfer> GetPossibleTransfers(IEnumerable<Stock> stocks)
		{
			var movings = new List<Transfer>();
			var stocksArray = stocks.ToArray();

			//Составляем список возможный перемещений
			foreach (var stockFrom in stocksArray)
			{
				foreach (var stockTo in stocksArray)
				{
					// Не составляем пару с одтнаковыми складами
					if (stockFrom != stockTo)
					{
						var moving = new Transfer { StockFrom = stockFrom, StockTo = stockTo };
						movings.Add(moving);
					}
				}
			}
			return movings;
		}

		// Возвращает из заданного списка ЗЧ, те позиции которые необходимо обеспечивать в наличии хотябы на одном складе
		public static List<OrderRequiredItem> GetOrderRequiredItems(Dictionary<string, Item> items)
		{
			var orderRequiredItems = new List<OrderRequiredItem>();
			foreach (var item in items)
			{
				var requiredStocks = new List<Stock>();
				var requiredItem = new OrderRequiredItem();
				// Если ЗЧ имеет RequiredAvailability, добавляем ее в список
				if (item.Value.RequiredAvailability)
				{
					requiredItem.Item = item.Value;

				}
				// Перебираем склады для ЗЧ
				foreach (var stock in item.Value.Stocks)
				{
					// Если склад имеет больше 3х продаж, добавляем данную зч в список для заказа
					if (stock.CountSelings >= 3 && requiredItem.Item != null)
					{
						requiredStocks.Add(stock);
					}
				}
				// Если складов с обязательным налчием нет или RequiredAvailability не была установлена, переходим к следующей ЗЧ
				if (!requiredStocks.Any() && requiredItem.Item == null) continue;
				requiredItem.OrderRequiredStocks = requiredStocks;
				orderRequiredItems.Add(requiredItem);
			}
			return orderRequiredItems;
		}

		// Формирует заказы
		public static List<Order> GetOrders(List<Order> orders, Dictionary<string, Item> items)
		{
			// Перебираем список ЗЧ
			foreach (var item in items)
			{
				// Определяем суммарный мин остаток и суммарный остаток
				var sumMinStocks = item.Value.GetSumMinStocks();
				var sumMaxStocks = item.Value.GetSumMaxStocks();
				var sumStocks = item.Value.GetSumStocks(false);
				var sumSelingKits = item.Value.GetSumSelings(true);

				// Если продалось меньше 3х комплектов, переходим к следующей зч
				// TODO временно, для тестов
				//if (sumSelingKits < 3)
				//{
				//	continue;
				//}

				// Делаем заказ только для запчастей имеющих директиву RequiredAvailability
				if (!item.Value.RequiredAvailability)
				{
					continue;
				}

				// Если общий остаток меньше общего минимального остатка и максимального, делаем заказ
				if (sumStocks <= sumMinStocks && sumStocks < sumMaxStocks)
				{
					var order = new Order
					{
						Item = item.Value,
						Count = Math.Floor((sumMaxStocks - sumStocks) / item.Value.InBundle) * item.Value.InBundle
					};
					orders.Add(order);
				}
				// Иначе добавляем в заказ с нулевым количеством
				else
				{
					var order = new Order
					{
						Item = item.Value,
						Count = 0
					};
					orders.Add(order);
				}
			}
			return orders;
		}

		// Создает переоценку
		public static List<Revaluation> GetRevaluations(Dictionary<string, Item> items)
		{
			var revaluations = new List<Revaluation>();

			// Перебираем ЗЧ
			foreach (var item in items)
			{
				// Если ЗЧ не имеет конкурентов, берем следующую ЗЧ
				if (item.Value.Сompetitors.Count == 0)
				{
					continue;
				}
				// Если остатка нет, берем следующую
				if (item.Value.GetSumStocks() == 0)
				{
					continue;
				}

				//// Переоценка
				var competitor = new Сompetitor();
				var note = "Не прописано правило";
				var allowSellingLoss = false;
				var withCopmetitorsStock = false;
				var deliveryTime = 7;
				var withDeliveryTime = true;

				// Если производитель "Китай"
				if (item.Value.Manufacturer == "Китай")
				{
					switch (item.Value.Property)
					{
						case "Норма":
						case "НП":
							withCopmetitorsStock = true;
							note = "Новый приход, Норма (остатки, срок доставки(7))";
							allowSellingLoss = false;
							deliveryTime = 7;
							withDeliveryTime = true;
							break;
						case "БП 2 мес":
						case "НЛ 12":
						case "НЛ 24":
							withCopmetitorsStock = false;
							withDeliveryTime = false;
							note = "БП 2 мес, НЛ 24, НЛ 12, (в минус)";
							allowSellingLoss = true;
							break;
						case "БП 1 мес":
							withCopmetitorsStock = false;
							withDeliveryTime = false;
							note = "БП 1 мес";
							allowSellingLoss = false;
							break;
						case "ОС 2":
						case "ОС 3":
							withCopmetitorsStock = true;
							withDeliveryTime = true;
							note = "ОС 2, ОС 3";
							allowSellingLoss = false;
							break;
					}
				}
				else
				{
					switch (item.Value.StorageCategory)
					{
						case "Попова":
						case "Везде":
						case "Нигде":
						case "МинЗапас":
							withCopmetitorsStock = false;
							note = "Попова, Везде, Нигде, МинЗапас (срок доставки(7))";
							allowSellingLoss = false;
							withDeliveryTime = true;
							deliveryTime = 7;
							break;
						case "НЛ12":
						case "НЛ24":
							withCopmetitorsStock = false;
							note = "НЛ 12, НЛ 24 (в минус)";
							allowSellingLoss = true;
							withDeliveryTime = false;
							break;
					}
				}
				competitor = item.Value.GetСompetitor(withDeliveryTime, withCopmetitorsStock, true, deliveryTime);
				note += " " + item.Value.OverStockDaysForAllStocks;
				var revaluation = new Revaluation(competitor, item.Value, note, allowSellingLoss);
				revaluations.Add(revaluation);
			}
			return revaluations;
		}

		// Создает книгу с заданным именем, вставляет в нее нужные данные и сохраняет
		public static void MakeImpot1CBook(Range inputRange, string bookName, string folder)
		{
			// Создаем новую книгу
			var impot1CBook = Globals.ThisWorkbook.Application.Workbooks.Add();
			// Вставляем на первый лист необходимые данные
			inputRange.Copy(impot1CBook.Worksheets[1].Range["A2"]);
			// Отключаем предупреждение и обновление экрана
			Starting(false, false);
			// Сохраняем книгу
			var fullPath = Path.Combine(Globals.ThisWorkbook.Path, "..\\", folder, bookName);
			impot1CBook.SaveAs(fullPath, XlFileFormat.xlWorkbookNormal);
			impot1CBook.Close();

			Ending();
		}

		// Изменяет состояние Application
		public static void Starting(bool displayAlerts = true, bool screenUpdating = true)
		{
			// Не показывать предупреждения
			if (!displayAlerts)
			{
				Globals.ThisWorkbook.Application.DisplayAlerts = false;
			}
			// Не обновлять экран
			if (!screenUpdating)
			{
				Globals.ThisWorkbook.Application.ScreenUpdating = false;
			}
		}

		// Изменяет состояние Application
		public static void Ending()
		{
			Globals.ThisWorkbook.Application.DisplayAlerts = true;
			Globals.ThisWorkbook.Application.ScreenUpdating = true;
		}

		public static void ArchiveBooks(string inputFolder, string outputFolder)
		{
			// Считываем все файлы в папке с перемещениями
			var fullPath = Path.Combine(Globals.ThisWorkbook.Path, "..\\", inputFolder);
			var arrayOfBooks = Directory.GetFiles(fullPath);
			// Перемещаем все файлы в архив
			foreach (var transferBook in arrayOfBooks)
			{
				// Определяем новый путь
				var destFileName = Path.Combine(Globals.ThisWorkbook.Path, "..\\", outputFolder, Path.GetFileName(transferBook));
				// Если такой файл есть в конечной папке, удаляем его
				if (File.Exists(destFileName))
				{
					File.Delete(destFileName);
				}
				File.Move(transferBook, destFileName);
			}

		}

		// Преобразует логическое значение в инт
		public static int BoolToInt(bool inPut)
		{
			return inPut ? 1 : 0;
		}
	}
}
