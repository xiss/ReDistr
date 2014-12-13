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

		// Дает список возможных перемещений
		public static List<Moving> GetPossibleMovings(IEnumerable<Stock> stocks)
		{
			var movings = new List<Moving>();
			var stocksArray = stocks.ToArray();

			//Составляем список возможный перемещений
			foreach (var stockFrom in stocksArray)
			{
				foreach (var stockTo in stocksArray)
				{
					// Не составляем пару с одтнаковыми складами
					if (stockFrom.Signature != stockTo.Signature)
					{
						var moving = new Moving { StockFrom = stockFrom, StockTo = stockTo };
						movings.Add(moving);
					}
				}
			}
			return movings;
		}
	}
}
