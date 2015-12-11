﻿using System;
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

		// Свойство ЗЧ 2
		public string Property;

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

		// Себестоимость
		public double CostPrice = 0;

		// Стоимость
		public double Price = 0;

		// Колличество дней перезатарки сумарно по всем складам
		public double OverStockDaysForAllStocks;

		// Комментарий, почему установлена RequiredAvailability
		public string NoteRequiredAvailability;

		// Остатки на складах
		public List<Stock> Stocks = new List<Stock>();

		// Конкуренты в Питерплюсе
		public List<Сompetitor> Сompetitors = new List<Сompetitor>();

		// Возвращает список всех возможных доноров, отсортированный по убыванию. Если задан список перемещений, то доноры выдаются из этого списка
		public List<Stock> GetListOfPossibleDonors(List<Transfer> existTransfers = null)
		{
			var listOfPossibleDonors = new List<Stock>();
			// Если список не задан, выдаем всех возможных доноров
			if (existTransfers == null)
			{
				// Если свободный осток отличен от нуля, то склад донор
				listOfPossibleDonors =
					Stocks.Where(stock => stock.FreeStock > 0).OrderByDescending(stock => stock.FreeStock).ToList();
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
			// Если задан OneDonor, выдаем свободные остатки только для этого донора
			if (Config.OneDonor != null)
			{
				return Stocks.Where(stock => stock == Config.OneDonor).Sum(stock => stock.FreeStock);
			}

			// Если список не задан, выдаем сумму для всех складов
			if (existTransfers == null)
			{
				return Stocks.Sum(stock => stock.FreeStock);
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
				sumSelings = sumSelings / InKit;
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

		// Проверяет, имеет ли хоть один склад директиву RequiredAvailability True
		public bool IsRequiredAvailability()
		{
			return Stocks.Any(stock => stock.RequiredAvailability);
		}

		// Возвращает ближаещего конкурента с учетом исключений
		public Сompetitor GetСompetitor(bool withDeliveryTime, bool withCompetitorsStocks, bool withExcludes = true, int deliveryTime = 0)
		{
			// Расчитываем параметры остатков у конкуретов
			const double minStockDays = 25;
			const double maxStockDays = 40;
			var sailsPerDay = GetSumSelings() / Config.SellingPeriod;

			var sumStocks = GetSumStocks();
			var minPercentStock = 0.2;
			// Если продаж не было то считаем что оверсток равен году
			if (sailsPerDay == 0)
			{
				OverStockDaysForAllStocks = 365;
			}
			else
			{
				OverStockDaysForAllStocks = Convert.ToInt32(Math.Round(sumStocks / sailsPerDay));
			}
			// Если запас более указанного уменьшаем процент
			if (OverStockDaysForAllStocks > maxStockDays)
			{
				minPercentStock = (maxStockDays / OverStockDaysForAllStocks) * 0.2;
			}
			// Если запас менее указанного увеличиваем процент
			else
			{
				minPercentStock = 0.65;
			}
			// minPercentStock должен быть между 0 и 1
			if (minPercentStock > 1)
			{
				minPercentStock = 1;
			}
			if (minPercentStock < 0)
			{
				minPercentStock = 0;
			}

			Сompetitors = Сompetitors.OrderBy(competitor => competitor.PositionNumber).ToList();

			foreach (var competitor in Сompetitors)
			{
				// Проверяем список исключений если конкуреты из этого списка переходим к следующему
				if (Config.ListExcludeCompetitors.Contains(competitor.Id) & withExcludes)
				{
					continue;
				}

				// Проверяем срок поставки, если не соответствует переходим к следующему
				if (competitor.DeliveryTime > deliveryTime & withDeliveryTime)
				{
					continue;
				}

				// Проверяем запас, если он меньше необходимого переходим к следующему
				if ((competitor.Count / sumStocks) < minPercentStock & withCompetitorsStocks)
				{
					continue;
				}

				return competitor;
			}
			return null;
		}

		// Возвращает новую цену расчитанную опираясь на указанного конкурента
		public double GetNewPrice(Сompetitor сompetitor, bool allowSellingLoss)
		{
			// Расчитываем новую цену
			double newPrice = 0;
			// Если конкурент есть
			if (сompetitor != null)
			{
				if (Manufacturer != "Китай")
				{
					switch (StorageCategory)
					{
						case "Попова":
						case "Везде":
						case "Нигде":
						case "МинЗапас":
							newPrice = CostPrice * 1.4;
							break;
						default:
							newPrice = сompetitor.Price * 0.87;
							break;
					}
				}
				else
				{
					newPrice = сompetitor.Price * 0.87;
				}
			}
			// Если конкурента нет
			else
			{
				// Если производитель "Китай"
				if (Manufacturer == "Китай")
				{
					switch (Property)
					{
						case "Норма":
						case "НП":
							newPrice = CostPrice * 2;
							break;
						case "БП 2 мес":
							newPrice = CostPrice;
							break;
						case "БП 1 мес":
							newPrice = CostPrice * 0.9;
							break;
						case "НЛ 12":
							newPrice = CostPrice * 0.95;
							break;
						case "НЛ 24":
							newPrice = CostPrice * 0.7;
							break;
						case "ОС 2":
						case "ОС 3":
							newPrice = CostPrice;
							break;
					}
				}
				else
				{
					switch (StorageCategory)
					{
						case "Попова":
						case "Везде":
						case "Нигде":
						case "МинЗапас":
							newPrice = CostPrice * 1.4;
							break;
						case "НЛ12":
							newPrice = CostPrice * 0.95;
							break;
						case "НЛ24":
							newPrice = CostPrice * 0.7;
							break;
					}
				}

			}
			// Если новая цена ниже себестоимости, возвращаем себестоимость
			if (newPrice < CostPrice && !allowSellingLoss)
			{
				newPrice = CostPrice;
			}
			return Math.Round(newPrice);
		}
	}
}