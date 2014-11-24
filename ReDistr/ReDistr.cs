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
					// TODO Проверить расчеты
					stock.SailPersent = GetSailPersent(item.Value, stock);
					stock.MinStock = GetMinStock(item.Value, stock);
					stock.MaxStock = GetMaxStock(item.Value, stock);
					stock.FreeStock = GetFreeStock(item.Value, stock);
					stock.Need = GetNeed(item.Value, stock);
				}
			}
		}

		// Возвращает процент продаж для указанного склада
		private static double GetSailPersent(Item item, Stock curentStock)
		{
			var curentStockSails = curentStock.SelingsCount;
			var allSails = item.Stocks.Sum(stock => stock.SelingsCount);

			// Проверка на нулевые продажи
			if (allSails == 0)
			{
				return 0;
			}

			return curentStockSails / allSails;
		}

		// Возвращает минимальный остаток для указанного склада
		private static double GetMinStock(Item item, Stock curentStock)
		{
			var sailsPerDay = curentStock.SelingsCount / Config.SellingPeriod;
			var minStock = Math.Ceiling((sailsPerDay * curentStock.DefaultPeriodMinStock) / item.InKit) * item.InKit;

			return minStock;
		}

		// Возвращает максимальный остаток для указанного склада
		private static double GetMaxStock(Item item, Stock curentStock)
		{
			var sailsPerDay = curentStock.SelingsCount / Config.SellingPeriod;
			var maxStock = Math.Ceiling((sailsPerDay * curentStock.DefaultPeriodMaxStock) / item.InKit) * item.InKit;

			return maxStock;
		}

		// Возвращает свободный остаток для указанного склада, вычислять после расчета мин остатка
		private static double GetFreeStock(Item item, Stock curentStock)
		{
			double freeStock;
			// Если мин остаток отличен от нуля
			if (curentStock.MinStock > 0)
			{
				freeStock = curentStock.Count - item.InKit - curentStock.InReserve;
			}
			// Если мин остаток равен 0
			else
			{
				freeStock = curentStock.Count;
			}

			// Свободный остаток не может быть меньше нуля
			if (freeStock < 0)
			{
				freeStock = 0;
			}

			return freeStock;

		}

		// Возвращает потребность для указанного склада, расчитывать после вычисления процента продаж
		private static double GetNeed(Item item, Stock curentStock, bool ceilingToKit = true)
		{
			//TODO нужно учитывать резервы при подсчете процентного количества
			var allCount = item.Stocks.Sum(curentItem => curentItem.Count);
			// Если потребность в процентах меньше макс остатка используем его
			double realMaxStock;
			if ((allCount * curentStock.SailPersent) > curentStock.MaxStock)
			{
				realMaxStock = curentStock.MaxStock;
			}
			// иначе берем проценты
			else
			{
				realMaxStock = allCount * curentStock.SailPersent;
			}

			// Расчитываем потребность
			var need = realMaxStock - curentStock.Count;

			// Если нужно делаем кратным комплекту и округляем в большую сторону
			if (ceilingToKit)
			{
				need = Math.Ceiling(need / item.InKit) * item.InKit;
			}

			// Потребность не может быть отрицательной
			if (need < 0)
			{
				need = 0;
			}

			return need;
		}

		// Дает список возможных перемещений
		public static List<Moving> GetPossibleMovings(IEnumerable<Stock> stocks)
		{
			var movings = new List<Moving>();

			//Составляем список возможный перемещений
			foreach (var stockFrom in stocks)
			{
				foreach (var stockTo in stocks)
				{
					// Не составляем пару с одтнаковыми складами
					if (stockFrom.Signature != stockTo.Signature)
					{
						var moving = new Moving {StockFrom = stockFrom, StockTo = stockTo};
						movings.Add(moving);
					}
				}
			}
			return movings;
		}

	}
}
