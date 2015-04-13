using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace ReDistr
{
	// Класс запчасть
	public class Item
	{
		// Код 1С, уникален
		public string Id1C;

		// Артикул
		public string Article;

		// Категория хранения
		public string StorageCategory;

		// Название 
		public string Name;

		// Производитель 
		public string Manufacturer;

		// Поставщик
		public string Supplier = Config.DefaultSupplierName;

		// Количество товара в комплекте, не может быть равен 0, больше 0.
		public double InKit = 1;

		// Количество товара в упаковке
		public double InBundle = 1;

		// Остатки на складах 
		public List<Stock> Stocks = new List<Stock>();

		// Возвращает список всех возможных доноров, отсортированный по убыванию. Если задан список перемещений, то доноры выдаются из этого списка
		public List<Stock> GetListOfPossibleDonors(List<Transfer> existTransfers = null)
		{
			var listOfPossibleDonors = new List<Stock>();
			// Если список не задан, выдаем всех возможных доноров
			if (existTransfers == null)
			{
				// Если свободный осток отличен от нуля, то склад донор
				listOfPossibleDonors = Stocks.Where(stock => stock.FreeStock > 0).OrderByDescending(stock => stock.FreeStock).ToList();
			}
			// Если список задан выдаем доноров из него
			else
			{
				foreach (var stock in Stocks)
				{
					foreach (var transfer in existTransfers)
					{
						if (stock == transfer.StockFrom)
						{
							listOfPossibleDonors.Add(stock);
						}
					}
				}
			}

			// Если задана дериктива одного донора, то оставляем только его в списке
			if (Config.OneDonor != null)
			{
				for (var i = 0; i < listOfPossibleDonors.Count; i++)
				{
					if (listOfPossibleDonors[i] != Config.OneDonor)
					{
						listOfPossibleDonors.Remove(listOfPossibleDonors[i]);
						i--;
					}
				}
			}
			return listOfPossibleDonors;
		}

		// Возвращает сумму всех свободных остатков, если задан список перемещений то остатки берутся из доноров в этих перемещениях
		public double GetSumFreeStocks(List<Transfer> existTransfers = null)
		{
			// Если список не задан, выдаем сумму для всех складов
			if (existTransfers == null)
			{
				return Stocks.Sum(stock => stock.FreeStock);
			}

			// Если задан OneDonor, выдаем свободные остатки только для этого донора
			if (Config.OneDonor != null)
			{
				return Stocks.Where(stock => stock == Config.OneDonor).Sum(stock => stock.FreeStock);
			}

			// Если задан список доноров, выдаем сумму свободных остатков по этим донорам
			var existDonors = new List<Stock>();
			foreach (var stock in Stocks)
			{
				foreach (var transfer in existTransfers)
				{
					if (stock == transfer.StockFrom)
					{
						existDonors.Add(stock);
					}
				}
			}

			return existDonors.Sum(stock => stock.FreeStock);
		}

		// Возвращает общее количество ЗЧ без учета резервов
		public double GetSumStocks(bool withReserve = true)
		{
			double sumStocks;

			// Если нужно учитываем резервы
			if (withReserve)
			{
				sumStocks = Stocks.Sum(stock => stock.Count - stock.InReserve);
			}
			else
			{
				sumStocks = Stocks.Sum(stock => stock.Count);
			}
			
			if (sumStocks < 0)
			{
				sumStocks = 0;
			}
			return sumStocks;
		}

		// Возвращает общий минимальный остаток
		public double GetSumMinStocks()
		{
			var sumMinStocks = Stocks.Sum(stock => stock.MinStock);
			return sumMinStocks;
		}

		// Возвращает общий максимальный остаток
		public double GetSumMaxStocks()
		{
			var sumMaxStocks = Stocks.Sum(stock => stock.MaxStock);
			return sumMaxStocks;
		}

		// Возвращает сумму продаж
		public double GetSumSelings(bool inKits = false)
		{
			var sumSelings = Stocks.Where(stock => stock.CountSelings > 0).Sum(stock => stock.CountSelings);

			// Переводим в комплекты
			if (inKits)
			{
				sumSelings = sumSelings/InKit;
			}

			return sumSelings;
		}

		// Обновляет свободные остатки
		public void UpdateFreeStocks(string typeFreeStock)
		{
			foreach (var stock in Stocks)
			{
				stock.UpdateFreeStock(this, typeFreeStock);
			}
		}
	}
}
