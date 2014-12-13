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

		// Возвращает список всех возможных доноров
		public List<Stock> GetListOfPossibleDonors()
		{
			var possybleDonors = new List<Stock>();

			foreach (var stock in Stocks)
			{
				if (stock.FreeStock > 0)
				{
					possybleDonors.Add(stock);
				}
			}

			return possybleDonors;
		}
	}
}
