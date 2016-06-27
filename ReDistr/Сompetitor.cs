using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;

namespace ReDistr
{
	// Конкуренты на порталах
	public class Сompetitor
	{
		// ЗЧ
		public Item Item;

		// Срок поставки
		public double DeliveryTime;

		// Цена
		private double _price;
		public double PriceWithoutAdd
		{
			get
			{
				double price = 0;
				double ratio;
				// Вычисляем наценку конкурента
				// Проверяем есть ли у нас данные по прошлой цене в П+, если данных нет, используем шаблон
				if (Item.Сompetitors.Exists(competitor => competitor.Id == "Наш прайс"))
				{
					// Определяем реальную наценку конкурента
					double ourPrice = Item.Сompetitors.Find(competitor => competitor.Id == "Наш прайс")._price;
					ratio = ourPrice / Item.Price;
					// Проверка на максимум
					if (ratio > 1.14)
					{
						ratio = 1.14;
					}
					// Проверка на минимум
					if (ratio < 1.11)
					{
						ratio = 1.11;
					}
				}
				else
				{
					ratio = 1.13;
				}
				price = _price / ratio;

				return price;
			}
			set { _price = value; }
		}

		public double PriceWithAdd
		{
			get { return _price; }
		}

		// Код поставщика
		public string Id;

		// Номер строки на портале
		public double PositionNumber;

		// Регион
		public string Region;

		// Количество
		public double Count;
	}
}
