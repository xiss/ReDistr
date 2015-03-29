using System;
using System.Collections.Generic;
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

		// Создает необходимые перемещения для обеспечения одного комплекта на складах с ненулевым минсоатком
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
					// Определяем количество для перемещения, если перемещать ничего не нужно переходим к следующему складу
					var need = stock.GetNeedToSafety(item.Value);
					// TODO Нужно оставить только уникальные перемещения по донору
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
					if (stockFrom.Signature != stockTo.Signature)
					{
						var moving = new Transfer { StockFrom = stockFrom, StockTo = stockTo };
						movings.Add(moving);
					}
				}
			}
			return movings;
		}

		// Создает книгу с заданным именем
		public static void MakeImpot1CBook(Range inputRange, string bookName, string folder, Application excel)
		{
			// Создаем новую книгу
			var impot1CBook = excel.Application.Workbooks.Add();
			// Вставляем на первый лист необходимые данные
			inputRange.Copy(impot1CBook.Worksheets[1].Range["A2"]);
			// Сохраняем книгу
			excel.DisplayAlerts = false;
			impot1CBook.SaveAs(Config.PuthToThisWb + Config.FolderTransfers + bookName, XlFileFormat.xlWorkbookNormal);
			impot1CBook.Close();
			excel.DisplayAlerts = true;
		}

		// Изменяет состояние Application
		// TODO разобраться с этой функцией
		//		public static void Starting(bool DisplayAlerts = true, bool ScreenUpdating = true, bool DisplayPageBreaks = true)
		//		{
		//			Application.DisplayAlerts = DisplayAlerts;
		//
		//		}

		// Архивирует перемещения
		public static void ArchiveTransfers()
		{
			// Считываем все файлы в папке с перемещениями
			var arrayOfTransferBooks = Directory.GetFiles(Config.PuthToThisWb + Config.FolderTransfers);
			// Перемещаем все файлы в архив
			foreach (var transferBook in arrayOfTransferBooks)
			{
				// Определяем новый путь
				var destFileName = Config.PuthToThisWb + Config.FolderArchiveTransfers + Path.GetFileName(transferBook);
				// Если такой файл есть в конечной папке, удаляем его
				if (File.Exists(destFileName))
				{
					File.Delete(destFileName);
				}
				File.Move(transferBook, destFileName);
			}

		}
	}
}
