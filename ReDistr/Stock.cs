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

		// Текущий остаток
		public double Count = 0;

		// Количество продаж из отчета
		public double CountSelings = 0;

		// Остаток до работы скрипта
		public double CountOrigin = 0;

		// Минимальный остаток, кратен inKit
		public double MinStock = 0;

		// Период для расчета мин остатка по умолчанию
		public uint DefaultPeriodMinStock;

		// Период для расчета макс остатка по умолчанию
		public uint DefaultPeriodMaxStock;

		// Максимальный остаток, кратен inKit
		public double MaxStock = 0;

		// Процент продаж
		public double SailPersent = 0;

		// Резерв, количество товара в резерве, всегда положителен
		public double InReserve = 0;

		// Приоритет, в спорных ситуациях используется для определения реципиента. 
		public uint Priority;

		// Себестоимость
		public double CostPrice = 0;

		// Список основных производителей для поддержания остатка
		public List<string> MainManufacturers;

		// Свободный остаток, который можно перемещать с данного склада, 
		// если свободный остаток отличен от 0, склад может быть донором. 
		// Не может быть меньше 0, если minStock отличен от нуля
		public double FreeStock;

		// Потребность данного склада в товаре, всегда положительна, 
		// если need отличен от 0, склад нуждается в товаре и может быть реципиентом. 
		// TODO Он вроде не нужен, проверить, везде используются спец функции
		public double Need;

		// Исключить данный склад из распределения
		// Выставляет свободный остатко равным реальному остатку за вычетом резервов
		public bool ExcludeFromMoovings;

		// Сигнатура склада, в нижнем регистре
		private string _signature;
		public string Signature
		{
			get { return _signature.ToLower(); }
			set { _signature = value; }
		}

		// Признак обязательного наличия ЗЧ на данном складе
		//public bool RequiredAvailability;

		// Перегруженный оператор ==
		public static bool operator ==(Stock a, Stock b)
		{
			// Оба могут быть нулом
			if (ReferenceEquals(a, b))
			{
				return true;
			}

			// Проверяем что не один не является нулом
			if (((object)a == null) || ((object)b == null))
			{
				return false;
			}

			// Проверяем что сигнатура одна и та же
			return a.Signature == b.Signature;
		}

		// Перегруженный оператор !=
		public static bool operator !=(Stock a, Stock b)
		{
			return !(a == b);
		}

		// Переопределенный оператор Equals
		public override bool Equals(Object obj)
		{
			if (obj.GetType() != typeof(Stock))
			{
				return false;
			}
			return this == (Stock)obj;
		}

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

			// Если мин остаток больше остатка, возвращаем потребность, иначе возвращаем 0
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
			var need = Math.Floor(Math.Abs(sumStocks * (MaxStock / sumStocks)) / item.InKit) * item.InKit - Count;

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
			var stockSails = CountSelings;
			var allSails = item.Stocks.Where(stock => stock.CountSelings > 0).Sum(stock => stock.CountSelings);

			// Проверка на нулевые продажи
			if (allSails == 0)
			{
				SailPersent = 0;
				return;
			}
			// Проверка на отрицательные продажи (баг в отчете)
			if (CountSelings < 0)
			{
				SailPersent = 0;
				return;
			}

			SailPersent = (stockSails / allSails);
		}

		// Расчитывает минимальный остаток для склада
		public void UpdateMinStock(Item item)
		{
			var sailsPerDay = CountSelings / Config.SellingPeriod;
			var minStock = Math.Ceiling((sailsPerDay * DefaultPeriodMinStock) / item.InKit) * item.InKit;

			// Если установлена RequiredAvailability и расчитанный минимальный остаток меньше кратности, проставлем кратность остаток равным одному комплекту
			if (item.RequiredAvailability && minStock < item.InKit)
			{
				minStock = item.InKit;
			}

			MinStock = minStock;
		}

		// Расчитывает максимальный остаток для указанного склада
		// TODO учитывать категории для расчета остатков, подумать нужно ли это делать
		public void UpdateMaxStock(Item item)
		{
			var sailsPerDay = CountSelings / Config.SellingPeriod;
			var maxStock = Math.Ceiling((sailsPerDay * DefaultPeriodMaxStock) / item.InKit) * item.InKit;

			// Если установлена RequiredAvailability и расчитанный максимальный остаток меньше кратности, проставлем остаток равным одному комплекту
			if (item.RequiredAvailability && maxStock < item.InKit)
			{
				maxStock = item.InKit;
			}

			MaxStock = maxStock;
		}

		// Расчитывает максимальный остаток с учетом основнх брендов, расчитывать после расчета максимальных остатков по уходимости
		// Переносит излишки на склад бренда
		public void UpdateMaxStockWithMainManufacturer(Item item)
		{
			if (MainManufacturers == null) return;
			if (!MainManufacturers.Contains(item.Manufacturer)) return;

			var newMaxStock = item.GetSumStocks(false) - item.GetSumMaxStocks();
			if (newMaxStock > MaxStock)
			{
				MaxStock = MaxStock + newMaxStock;
			}
			// Если мин остаток 0 то ставим еденицу чтобы спровоцировать перемещение
			if (MinStock == 0)
			{
				MinStock = 1;
			}
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

		// Устанавливает начальные остатки
		public void SetOriginCount(double count)
		{
			Count = count;
			CountOrigin = count;
		}
	}
}
