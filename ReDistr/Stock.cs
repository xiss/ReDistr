using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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

		// Резерв, количество товара в резерве
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

		// Возвращает требуемое количество ЗЧ для удовлетворения кратности
		public double GetNeedToInKit(Item item)
		{
			double need = 0;

			// Если мин остаток больше нуля и остаток на складе меньше кратности, возвращаем потребность, иначе возвращаем 0
			if (MinStock > 0 && Count > item.InKit)
			{
				need = item.InKit - Count;
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
		public void UpdateFreeStock(Item item)
		{
			double freeStock;
			// Если мин остаток отличен от нуля
			if (MinStock > 0)
			{
				freeStock = Count - item.InKit - InReserve;
			}
			// Если мин остаток равен 0
			else
			{
				freeStock = Count;
			}

			// Свободный остаток не может быть меньше нуля
			if (freeStock < 0)
			{
				freeStock = 0;
			}

			FreeStock = freeStock;
		}

		// Расчитывает потребность для указанного склада, расчитывать после вычисления процента продаж
		public void UpdateNeed(Item item, bool ceilingToKit = true)
		{
			//TODO нужно учитывать резервы при подсчете процентного количества
			var allCount = item.Stocks.Sum(curentItem => curentItem.Count);
			// Если потребность в процентах меньше макс остатка используем его
			double realMaxStock;
			if ((allCount * SailPersent) > MaxStock)
			{
				realMaxStock = MaxStock;
			}
			// иначе берем проценты
			else
			{
				realMaxStock = allCount * SailPersent;
			}

			// Расчитываем потребность
			var need = realMaxStock - Count;

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

			Need = need;
		}
	}
}
