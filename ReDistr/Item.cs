using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
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

		// Количество товара в комплекте, не может быть равен 0, больше 0.
		public double InKit = 1;

		// Количество товара в упаковке
		public double InBundle = 1;

		// Остатки на складах 
		public List<Stock> Stocks = new List<Stock>();

		// Возвращает список всех возможных доноров, отсортированный по убыванию
		public List<Stock> GetListOfPossibleDonors()
		{
			// Если свободный осток отличен от нуля, то склад донор
			return Stocks.Where(stock => stock.FreeStock > 0).OrderByDescending(stock => stock.FreeStock).ToList();
		}

		// Возвращает сумму всех свободных остатков
		public double GetSumFreeStock()
		{
			return Stocks.Sum(stock => stock.FreeStock);
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
