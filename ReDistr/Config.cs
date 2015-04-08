using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReDistr
{
	static class  Config
	{
		// Дата снятия отчета с остатками
		public static DateTime StockDate;

		// Дата начала периода продаж
		public static DateTime PeriodSellingFrom;

		// Дата окончания периода продаж
		public static DateTime PeriodSellingTo;

		// Количество дней в периоде продаж
		public static int SellingPeriod;

		// Имя книги с остатками
		public static string NameOfStocksWb;

		// Имя книги с продажами
		public static string NameOfSealingsWb;

		// Папка с перемещениями
		public static string FolderTransfers = "Перемещения\\";

		// Папка с архивом перемещений
		public static string FolderArchiveTransfers = "Перемещения\\Архив\\";

		// Имя книги с параметрами
		public static string NameOfParametersWb;

		// Количество складов
		public static uint StockCount;

		// Количество возможных перемещений
		public static int CountPossibleTransfers;

		// Список возможных перемещений
		public static List<Transfer> PossibleTransfers;

		// Показывать отчет со всеми ЗЧ
		// TODO /1 учитывать
		public static bool ShowReport = true;

		// Минимальное количество проданных комплектов для расчет мин остатка
		// TODO /5 учитывать
		public static double MinSoldKits = 0;

		// Если параметр указан, то перемещения делать только с этого склада
		public static Stock OneDonor = null;

		// Категории хранения товара для перемещения
		public static List<string> ListStorageCategoryToTransfers;

		// Устанавливает список возможных перемещений и их количество
		public static void SetPossibleTransfers()
		{
			PossibleTransfers = ReDistr.GetPossibleTransfers(SimpleStockFactory.CurrentFactory.GetAllStocks()).ToList();
			CountPossibleTransfers = PossibleTransfers.Count;
		}
			

	}
}
