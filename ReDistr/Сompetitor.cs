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
			    if (_price > 0 & _price < 7)
			    {
			        price = _price / 1.148;
			    }
                else if (_price > 7 & _price < 1136)
			    {
                    price = _price / 1.147;
			    }
                else if (_price > 1136 & _price < 2222)
                {
                    price = _price / 1.140;
                }
                else if (_price > 2222 & _price < 4473)
                {
                    price = _price / 1.133;
                }
                else if (_price > 4473 & _price < 5570)
                {
                    price = _price / 1.126;
                }
                else if (_price > 5570 & _price < 6673)
                {
                    price = _price / 1.122;
                }
                else if (_price > 6672 & _price < 7770)
                {
                    price = _price / 1.118;
                }
                else if (_price > 7770 & _price < 8860)
                {
                    price = _price / 1.115;
                }
                else if (_price > 8860 & _price < 16663)
                {
                    price = _price / 1.111;
                }
                else if (_price > 16663 & _price < 100000000)
                {
                    price = _price / 1.104;
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
