using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReDistr
{
	// Класс перемещение
	class Moving
	{
		// Откуда перемещение
		public Stock StockFrom;

		// Куда перемещение
		public Stock StockTo;

		// Список ЗЧ
		// TODO Как указывать количество ЗЧ?
		public Dictionary<string, Item> Items;
	}
}
