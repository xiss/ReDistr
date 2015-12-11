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
			NewPrice = item.GetNewPrice(competitor, allowSellingLoss);
			Item = item;

			// Если нет конкурента
			if (competitor == null)
			{
				Note = "Нет конкурента";
			}
			// Если конкурент есть
			else
			{
				Note = note;
				Competitor = competitor;
			}
		}
	}
}
