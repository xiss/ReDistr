﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReDistr
{
	// Класс перемещение
	class Transfer
	{
		// Откуда перемещение
		public Stock StockFrom;

		// Куда перемещение
		public Stock StockTo;

		// Колличество ЗЧ
		public double Count;
		
		// ЗЧ
		public Item Item;

		// Применяет перемещение, обновляет остатки в соответствии с ними
		public void Apply()
		{
			StockFrom.Count -= Count;
			StockFrom.FreeStock -= Count;
			StockTo.Count += Count;
		}

		//TODO создаить метод формирующий перемещения (на вход итемы на выход список трансферов)
	}
}