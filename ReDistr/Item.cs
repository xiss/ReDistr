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
		public string Supplier;

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
			if (Config.OneDonorSignature != null)
			{
				for (var i = 0; i < listOfPossibleDonors.Count; i++)
				{
					if (!listOfPossibleDonors[i].Signature.ToLower().Contains(Config.OneDonorSignature.ToLower()))
					{
						listOfPossibleDonors.Remove(listOfPossibleDonors[i]);
						i--;
					}
				}
			}
			return listOfPossibleDonors;
		}

		// Возвращает сумму всех свободных остатков, если задан список перемещений то остатки берутся из доноров в этих перемещениях
		//  TODO /10 учитывать OneDonorSignature
		public double GetSumFreeStocks(List<Transfer> existTransfers = null)
		{
			// Если список не задан, выдаем сумму для всех складов
			if (existTransfers == null)
			{
				return Stocks.Sum(stock => stock.FreeStock);
			}

			var ExistDonors = new List<Stock>();
			foreach (var stock in Stocks)
			{
				foreach (var transfer in existTransfers)
				{
					if (stock == transfer.StockFrom)
					{
						ExistDonors.Add(stock);
					}
				}
			}

			return ExistDonors.Sum(stock => stock.FreeStock);
		}

		// Возвращает общее количество ЗЧ без учета резервов
		public double GetSumStocks()
		{
			var sumStocks = Stocks.Sum(stock => stock.Count - stock.InReserve);
			return sumStocks;
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
