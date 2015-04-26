using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReDistr
{
	public class OrderRequiredItem
	{
		// Список складов для обазятельного наличия данной ЗЧ
		public List<Stock> OrderRequiredStocks;

		// ЗЧ
		public Item Item;
	}
}
