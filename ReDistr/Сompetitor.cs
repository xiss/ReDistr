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
				// double ratio;
                // Вариант с пристрелкой
				// Вычисляем наценку конкурента
				// Проверяем есть ли у нас данные по прошлой цене в П+, если данных нет, используем шаблон
                //if (Item.Сompetitors.Exists(competitor => competitor.Id == "Наш прайс"))
                //{
                //    // Определяем реальную наценку конкурента
                //    double ourPrice = Item.Сompetitors.Find(competitor => competitor.Id == "Наш прайс")._price;
                //    ratio = ourPrice / Item.Price;
                //    // Проверка на максимум
                //    if (ratio > 1.14)
                //    {
                //        ratio = 1.14;
                //    }
                //    // Проверка на минимум
                //    if (ratio < 1.11)
                //    {
                //        ratio = 1.11;
                //    }
                //}
                //else
                //{
                //    ratio = 1.13;
                //}
                //price = _price / ratio;

                // Вариант с порогами
			    if (price > 0 & price < 6)
			    {
			        price = price * 0.852;
			    }
                else if (price > 7 & price < 990)
			    {
                    price = price * 0.853;
			    }
                else if (price > 991 & price < 1950)
                {
                    price = price * 0.860;
                }
                else if (price > 1950 & price < 3950)
                {
                    price = price * 0.867;
                }
                else if (price > 3951 & price < 4950)
                {
                    price = price * 0.874;
                }
                else if (price > 4951 & price < 5950)
                {
                    price = price * 0.878;
                }
                else if (price > 5951 & price < 6950)
                {
                    price = price * 0.882;
                }
                else if (price > 6951 & price < 7950)
                {
                    price = price * 0.885;
                }
                else if (price > 7951 & price < 15000)
                {
                    price = price * 0.889;
                }
                else if (price > 15001 & price < 100000000)
                {
                    price = price * 0.9;
                }
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
