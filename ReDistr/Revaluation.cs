using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReDistr
{

	// Переоценка
	public class Revaluation
	{
		// Новая цена
		public double NewPrice;

		//ЗЧ
		public Item Item;

		// Выбранный конкурент
		public Сompetitor Competitor;

		// Отладочная информация
		public string Note;

		// Конструктор
		public Revaluation(Сompetitor competitor, Item item, String note, bool allowSellingLoss)
		{
			// Если нет конкурента
			if (competitor == null)
			{
				Item = item;
				Note = "Нет конкурента";
			}
			// Если конкурент есть
			else
			{
				Item = item;
				NewPrice = item.GetNewPrice(competitor, allowSellingLoss);
				Note = note;
				Competitor = competitor;
			}
		}
	}
}
