using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using System.Text;

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
					// TODO Проверить расчеты - вроде все верно
					stock.UpdateSailPersent(item.Value);
					stock.UpdateMinStock(item.Value);
					stock.UpdateMaxStock(item.Value);
					stock.UpdateFreeStock(item.Value);
					stock.UpdateNeed(item.Value);
				}
			}
		}

		// Создает необходимые перемещения для обеспечения одного комплекта на складах с ненулевым минсоатком
		public static List<Transfer> GetTransfers(Dictionary<string, Item> items)
		{
			var transfers = new List<Transfer>();

			// Перебираем список ЗЧ
			foreach (KeyValuePair<string, Item> item in items)
			{
				// TODO Хз работате ли, это сортировка по приоритету
				item.Value.Stocks.OrderBy(stock => stock.Priority);

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
					if (need > item.Value.GetSumFreeStock())
					{
						continue;
					}
					// Создаем необходимые перемещения от доноров
					foreach (var possibleDonor in possibleDonors)
					{
						// Если потребность удовлетворяется одним донором
						if (need < possibleDonor.FreeStock)
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
								Count = need,
								Item = item.Value
							};
							transfer.Apply();
							transfers.Add(transfer);
						}
					}
				}
			}
			return transfers;
		}
		// Дает список возможных перемещений
		public static List<Transfer> GetPossibleMovings(IEnumerable<Stock> stocks)
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
	}
}
