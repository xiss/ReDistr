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
			    if (_price > 0 & _price <= 1146.49)
			    {
			        price = _price / 1.1581;
			    }
                else if (_price >= 1146.50 & _price <= 2244.05)
                {
                    price = _price / 1.1508;
                }
                else if (_price >= 2244.06 & _price <= 4517.00)
                {
                    price = _price / 1.1435;
                }
                else if (_price >= 4517.01 & _price <= 5624.76)
                {
                    price = _price / 1.1363;
                }
                else if (_price >= 5624.77 & _price <= 6739.62)
                {
                    price = _price / 1.1327;
                }
                else if (_price >= 6739.63 & _price <= 7847.30)
                {
                    price = _price / 1.1291;
                }
                else if (_price >= 7847.31 & _price <= 8947.83)
                {
                    price = _price / 1.1255;
                }
                else if (_price >= 8947.84 & _price <= 16828.86)
                {
                    price = _price / 1.1219;
                }
                else if (_price >= 16663 & _price <= 100000000)
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
