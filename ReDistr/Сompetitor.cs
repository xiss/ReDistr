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

		// Срок поставки
		public double DeliveryTime;

		// Цена
		private double _price;
		public double Price
		{
			get
			{
				var oldPrice = _price;
				_price = _price / 0.87;
				if (oldPrice > 0 && oldPrice < 999)
				{
					_price = _price * 0.8764;
				}
				if (oldPrice > 1001 && oldPrice < 1999)
				{
					_price = _price * 0.8817;
				}
				if (oldPrice > 2000 && oldPrice < 3999)
				{
					_price = _price * 0.8871;
				}
				if (oldPrice > 4000 && oldPrice < 5999)
				{
					_price = _price * 0.8953;
				}
				if (oldPrice > 6000 && oldPrice < 7999)
				{
					_price = _price * 0.8981;
				}
				if (oldPrice > 8000 && oldPrice < 99999999)
				{
					_price = _price * 0.9036;
				}
				return _price;
			}
			set { _price = value; }
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
