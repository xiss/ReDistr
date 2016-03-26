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
		// public double GetAVGCostPrice() = 0;

		// Стоимость
		public double Price = 0;

		// Колличество дней перезатарки сумарно по всем складам
		public double OverStockDaysForAllStocks;

		// Комментарий, почему установлена RequiredAvailability
		// public string NoteRequiredAvailability;

		// Остатки на складах
		public List<Stock> Stocks = new List<Stock>();

		// Конкуренты в Питерплюсе
		public List<Сompetitor> Сompetitors = new List<Сompetitor>();

		// Предустановленная цена
		public double PrePrice = 0;

		// Признак обязательного наличия ЗЧ на данном складе
		public bool RequiredAvailability;

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

		// Возвращает среднюю себестоимость
		public double GetAVGCostPrice()
		{
			var a = Stocks.Sum(stock => stock.CostPrice);
			var b = GetSumStocks(false);
			var c = Math.Round(Stocks.Sum(stock => stock.CostPrice * stock.Count) / GetSumStocks(false), 2);
			return Math.Round(Stocks.Sum(stock => stock.CostPrice * stock.Count) / GetSumStocks(false), 2);
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
		//		public bool IsRequiredAvailability()
		//		{
		//			return Stocks.Any(stock => stock.RequiredAvailability);
		//		}

		// Возвращает ближаещего конкурента с учетом исключений
		public Сompetitor GetСompetitor(bool withDeliveryTime, bool withCompetitorsStocks, bool withExcludes = true, int deliveryTime = 0)
		{
			var sumStocks = GetSumStocks();

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
				if (competitor.Count < sumStocks / 10 & withCompetitorsStocks)
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

			// Если есть предустановленная цена, используем ее
			if (PrePrice != 0)
			{
				return PrePrice;
			}

			// Если конкурент есть
			if (сompetitor != null)
			{
				//Если не Китай
				if (Manufacturer != "Китай")
				{
					switch (StorageCategory)
					{
						case "Попова":
						case "Везде":
						case "Нигде":
						case "МинЗапас":
							if (сompetitor.Price < GetAVGCostPrice() * 1.4)
							{
								newPrice = GetAVGCostPrice() * 1.4;
							}
							else
							{
								newPrice = сompetitor.Price;
							}
							break;
						case "НЛ12":
							if (сompetitor.Price > GetAVGCostPrice() * 0.95)
							{
								newPrice = GetAVGCostPrice() * 0.95;
							}
							else
							{
								newPrice = сompetitor.Price;
							}
							break;
						case "НЛ24":
							if (сompetitor.Price > GetAVGCostPrice() * 0.7)
							{
								newPrice = GetAVGCostPrice() * 0.7;
							}
							else
							{
								newPrice = сompetitor.Price;
							}
							break;
						default:
							newPrice = сompetitor.Price;
							break;
					}
				}
				//Если Китай
				else
				{
					switch (Property)
					{
						case "НЛ 12":
							if (сompetitor.Price > GetAVGCostPrice() * 0.95)
							{
								newPrice = GetAVGCostPrice() * 0.95;
							}
							else
							{
								newPrice = сompetitor.Price;
							}
							break;
						case "НЛ 24":
							if (сompetitor.Price > GetAVGCostPrice() * 0.7)
							{
								newPrice = GetAVGCostPrice() * 0.7;
							}
							else
							{
								newPrice = сompetitor.Price;
							}
							break;
						case "БП 1 мес":
							newPrice = (сompetitor.Price);
							break;
						case "БП 2 мес":
							newPrice = (сompetitor.Price);
							break;
						case "ОС 2":
							newPrice = сompetitor.Price;
							break;
						default:
							newPrice = сompetitor.Price;
							break;
					}
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
							newPrice = GetAVGCostPrice() * 2;
							break;
						case "БП 2 мес":
							newPrice = GetAVGCostPrice() * 0.8;
							break;
						case "БП 1 мес":
							newPrice = GetAVGCostPrice() * 1.1;
							break;
						case "НЛ 12":
							newPrice = GetAVGCostPrice() * 0.95;
							break;
						case "НЛ 24":
							newPrice = GetAVGCostPrice() * 0.7;
							break;
						case "ОС 2":
						case "ОС 3":
							newPrice = GetAVGCostPrice() * 2;
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
							newPrice = GetAVGCostPrice() * 1.4;
							break;
						case "НЛ12":
							newPrice = GetAVGCostPrice() * 0.95;
							break;
						case "НЛ24":
							newPrice = GetAVGCostPrice() * 0.7;
							break;
					}
				}
			}
			// Если новая цена ниже себестоимости, возвращаем себестоимость
			if (newPrice < (GetAVGCostPrice() * 1.1) && !allowSellingLoss)
			{
				newPrice = GetAVGCostPrice() * 1.1;
			}
			return Math.Round(newPrice, 2);
		}
	}
}