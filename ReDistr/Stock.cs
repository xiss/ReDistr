using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms.VisualStyles;

namespace ReDistr
{
	// Класс склад
	public class Stock
	{
		// Название склада
		public string Name;

		// Остаток  до выполнения перемещений
		public double Count;

		// Количество продаж из отчета
		public double SelingsCount;

		// Остаток после выполнения перемещений
		public uint InStockBefore;

		// Минимальный остаток, кратен inKit
		public double MinStock;

		// Период для расчета мин остатка по умолчанию
		public uint DefaultPeriodMinStock;

		// Период для расчета макс остатка по умолчанию
		public uint DefaultPeriodMaxStock;

		// Максимальный остаток, кратен inKit
		public double MaxStock;

		// Процент продаж
		public double SailPersent;

		// Резерв, количество товара в резерве, всегда положителен
		public double InReserve;

		// Приоритет, в спорных ситуациях используется для определения реципиента. 
		public uint Priority;

		// Свободный остаток, который можно перемещать с данного склада, 
		// если свободный остаток отличен от 0, склад может быть донором. 
		// Не может быть меньше 0, если minStock отличен от нуля
		public double FreeStock;

		// Потребность данного склада в товаре, всегда положительна, 
		// если need отличен от 0, склад нуждается в товаре и может быть реципиентом. 
		public double Need;

		// Исключить данный склад из распределения
		// Выставляет свободный остатко равным реальному остатку за вычетом резервов
		public bool ExcludeFromMoovings;

		// Сигнатура склада
		public string Signature;

		// Возвращает требуемое количество ЗЧ для обеспечения одного комплекта на площадке
		public double GetNeedToInKit(Item item)
		{
			double need = 0;

			// Если мин остаток больше нуля и остаток на складе меньше кратности, возвращаем потребность, иначе возвращаем 0
			if (MinStock > 0 && Count < item.InKit)
			{
				need = item.InKit - Count;
			}

			return need;
		}

		// Возвращает требуемое количество ЗЧ для обеспечения мин. остатка
		public double GetNeedToMinStock(Item item)
		{
			double need = 0;

			// Если мин остаток меньше остатка, возвращаем потребность, иначе возвращаем 0
			if (MinStock > Count)
			{
				need = MinStock - Count;
			}

			return need;
		}

		// Возвращает требуемое количество ЗЧ для обеспечения запаса
		public double GetNeedToSafety(Item item)
		{
			// Получаем общее колличество ЗЧ без учета резервов
			var sumStocks = item.GetSumStocks();
			var need = Math.Ceiling(Math.Abs(sumStocks * SailPersent) / item.InKit) * item.InKit - Count;

			// Итоговое количество не должно быть больше максимального остатка
			if ((need + Count) > MaxStock)
			{
				need = MaxStock - Count;
			}

			// Если потребность мееньше 0, округляем ее до 0
			if (need < 0)
			{
				need = 0;
			}

			return need;
		}

		// Определяет, удовлетворен ли мин остаток
		public bool NeedToMinStock()
		{
			// Если количество на складе меньше либо равно мин остатку, мин остаток не удовлетворен
			if (Count <= MinStock)
			{
				return true;
			}

			return false;
		}

		// Расчитывает процент продаж для склада
		public void UpdateSailPersent(Item item)
		{
			var stockSails = SelingsCount;
			var allSails = item.Stocks.Sum(stock => stock.SelingsCount);

			// Проверка на нулевые продажи
			if (allSails == 0)
			{
				SailPersent = 0;
				return;
			}

			SailPersent = (stockSails / allSails);
		}

		// Расчитывает минимальный остаток для склада
		public void UpdateMinStock(Item item)
		{
			var sailsPerDay = SelingsCount / Config.SellingPeriod;
			var minStock = Math.Ceiling((sailsPerDay * DefaultPeriodMinStock) / item.InKit) * item.InKit;

			MinStock = minStock;
		}

		// Расчитывает максимальный остаток для указанного склада
		public void UpdateMaxStock(Item item)
		{
			var sailsPerDay = SelingsCount / Config.SellingPeriod;
			var maxStock = Math.Ceiling((sailsPerDay * DefaultPeriodMaxStock) / item.InKit) * item.InKit;

			MaxStock = maxStock;
		}

		// Расчитывает свободный остаток для указанного склада, вычислять после расчета мин остатка
		public void UpdateFreeStock(Item item, string typeFreeStock)
		{
			double freeStock = 0;
			switch (typeFreeStock)
			{
				case "kit":
					// Если мин остаток отличен от нуля
					if (MinStock > 0)
					{
						freeStock = Count - item.InKit - InReserve;
					}
					// Если мин остаток равен 0
					else
					{
						freeStock = Count - InReserve;
					}
					break;

				case "minStock":
					// Если мин остаток отличен от нуля
					if (MinStock > 0)
					{
						freeStock = Count - MinStock - InReserve;
					}
					// Если мин остаток равен 0
					else
					{
						freeStock = Count - InReserve;
					}
					break;
			}

			// Свободный остаток не может быть меньше нуля
			if (freeStock < 0)
			{
				freeStock = 0;
			}

			FreeStock = freeStock;
		}
	}
}
