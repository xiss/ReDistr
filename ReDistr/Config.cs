using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

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

		// Имя книги с конкурентами
		public static string NameOfCompetitorsWb;

		// id нашего прайса на П+
		public static string IdPriceAp;

		// Папка с перемещениями
		public static string FolderTransfers = "Перемещения\\";

		// Папка с архивом перемещений
		public static string FolderArchiveTransfers = "Перемещения\\Архив\\";

		// Папка с переоценками
		public static string FolderRevaluations = "Переоценки\\";

		// Папка с архивом переоценками
		public static string FolderArchiveRevaluations = "Переоценки\\Архив\\";

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

		// Склад для перемещения выбранных категорий (неликвид)
		public static Stock StockToTransferSelectedStorageCategory = null;

		// Список категорий для перемещения на выбранный склад
		public static List<String> ListSelectedStorageCategoryToTransfer;

		// Если параметр указан, то перемещения делать только с этого склада
		public static Stock OneDonor = null;

		// Склад для оптовых отгрузок
		public static Stock WholesaleStock;

		// Категории хранения товара для перемещения
		public static List<string> ListStorageCategoryToTransfers;

		// Список конкурентов исключений
		public static List<string> ListExcludeCompetitors;

		// Список поставщиков
		public static List<string> ListSuppliers;

		// Значение параметра Supplier по умолчанию у ЗЧ
		public static string DefaultSupplierName = "none";

		// Выполнялся ли парс Остатков
		public static bool ParsedStocks = false;

		// Выполнялся ли парс Продаж
		public static bool ParsedSealings = false;

		// Выполнялся ли парс дополнительных параметров
		public static bool ParsedAdditionalParameters = false;

		// Выполнялся ли парс конкурентов
		public static bool ParsedCompetitors = false;

		// Процент остатка от нашего склада для рассмотрения его как конкурента
		public static double MinStockForCompetitor;

		// Устанавливает список возможных перемещений и их количество
		public static void SetPossibleTransfers()
		{
			PossibleTransfers = ReDistr.GetPossibleTransfers(SimpleStockFactory.CurrentFactory.GetAllStocks()).ToList();
			CountPossibleTransfers = PossibleTransfers.Count;
		}

		// ТЕСТ
		//public static string Test()
		//{
		//	return XmlSerializer.
		//}
			

	}
}
