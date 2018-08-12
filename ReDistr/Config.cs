using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.IO;
using System.Runtime.InteropServices;
using NLog;

namespace ReDistr.Config
{
	public class Config
	{
		public Files Files;
		public Revaluations Revaluations;
		public Orders Orders;
		public Transfers Transfers;

	    public List<Stock> Stocks;

		private static readonly string ConfigFile = AppDomain.CurrentDomain.BaseDirectory + "../config.xml";
		
		/// <summary>
		/// Загрузить настройки 
		/// </summary>
		protected static Config Load()
		{
			try
			{
				var serializer = new XmlSerializer(typeof(Config));
			    using (var stream = File.OpenRead(ConfigFile))
				{
					return (Config)serializer.Deserialize(stream);
				}
			}
			catch (Exception e)
			{
				LogManager.GetCurrentClassLogger().Error("Ошибка загрузки настроек. {0}", e.Message);
				// TODO как закончить выполнение функции, вызваться может где угодно
			}
		    return null;
		}
		/// <summary>
		/// Сохранить настройки
		/// </summary>
		public static void Save()
		{
			try
			{
				var serializer = new XmlSerializer(typeof(Config));
				Stream writer = new FileStream(ConfigFile, FileMode.Create);
				serializer.Serialize(writer, Inst);
				writer.Close();
			}
			catch (Exception e)
			{
				LogManager.GetCurrentClassLogger().Error("Ошибка сохранения настроек. {0}", e.Message);
			}
		}

		//Singleton
		private Config() { }
		private static Config _inst;
	    public static Config Inst => _inst;
		static Config()
		{
		    _inst = Load();
		}

        // Дата снятия отчета с остатками
		public static DateTime StockDate;

		// Склад для перемещения выбранных категорий (неликвид)
		public static Stock StockToTransferSelectedStorageCategory => SimpleStockFactory.CurrentFactory.GetStock(Inst.Transfers.StockNameToTransferSelectedStorageCategory);

		// Если параметр указан, то перемещения делать только с этого склада
		public static Stock OneDonor => SimpleStockFactory.CurrentFactory.GetStock(Inst.Transfers.StockNameOneDonor);

		// Склад для оптовых отгрузок
		public static Stock WholesaleStock => SimpleStockFactory.CurrentFactory.GetStock(Inst.Revaluations.StockNameWholesaleStock);

		// Дата начала периода продаж
		public static DateTime PeriodSellingFrom;

		// Дата окончания периода продаж
		public static DateTime PeriodSellingTo;

		// Количество дней в периоде продаж
		public static int SellingPeriod;

		// Количество складов
		public static uint StockCount => (uint)Inst.Stocks.Count;

		// Количество возможных перемещений
		public static int CountPossibleTransfers;

		// Список возможных перемещений
		public static List<Transfer> PossibleTransfers;

		// Минимальное количество проданных комплектов для расчет мин остатка
		// TODO /5 учитывать
		public static double MinSoldKits = 0;

		// Список поставщиков
		public static List<string> ListSuppliers;

		// Выполнялся ли парс Остатков
		public static bool ParsedStocks = false;

		// Выполнялся ли парс Продаж
		public static bool ParsedSealings = false;

		// Выполнялся ли парс дополнительных параметров
		public static bool ParsedAdditionalParameters = false;

		// Выполнялся ли парс конкурентов
		public static bool ParsedCompetitors = false;

		// Устанавливает список возможных перемещений и их количество
		public static void SetPossibleTransfers()
		{
			PossibleTransfers = ReDistr.GetPossibleTransfers(SimpleStockFactory.CurrentFactory.GetAllStocks()).ToList();
			CountPossibleTransfers = PossibleTransfers.Count;
		}

	}

	public class Files
	{
		// Папка с перемещениями
		public string FolderTransfers;

		// Папка с архивом перемещений
		public string FolderArchiveTransfers;

		// Папка с переоценками
		public string FolderRevaluations;

		// Папка с архивом переоценками
		public string FolderArchiveRevaluations;

		// Имя книги с остатками
		public string NameOfStocksWb;

		// Имя книги с продажами
		public string NameOfSealingsWb;

		// Имя книги с конкурентами
		public string NameOfCompetitorsWb;

		// Имя книги с параметрами
		public string NameOfParametersWb;
	}

	public class Transfers
	{
		// Категория обязательного наличия
		public string NameOfStorageCatRequiredAvailability;

		// имя склада для перемещения выбранных категорий (неликвид)
		public string StockNameToTransferSelectedStorageCategory;

		// Список категорий для перемещения на выбранный склад
		public List<String> ListSelectedStorageCategoryToTransfer;

		// Если параметр указан, то перемещения делать только с этого склада
		public string StockNameOneDonor = null;

		// Категории хранения товара для перемещения
		public List<string> ListStorageCategoryToTransfers;

		// Свойства ЗЧ для обязательного наличия, синоним мин. остатка (свойство ЗЧ1)
		public List<string> ListPropertyRequiredAvailability;
		
	}

	public class Revaluations
	{
		// id нашего прайса на П+
		public string IdPriceAp;

		// Склад для оптовых отгрузок
		public string StockNameWholesaleStock;

		// Список конкурентов исключений
		public List<string> ListExcludeCompetitors;

		// Процент остатка от нашего склада для рассмотрения его как конкурента
		public double MinStockForCompetitor;

		// Наш срок поставки
		public double OurDeliveryTime;

		// Процент для выявления демпинга
		public double DumpingPersent;

		// Допустимая разница в сроке доставки между нами и конкурентом
		public double DeltaDeliveryTime;

		// Отношение остатков конкурента к нашему
		public double DeltaCompetitorStock;

		// Максимально можно пропустить конкурентов
		public double MaxCompetitorsToMiss;

		// типа конкурента
		public int TypeCompetitor;

		// Коэффециент корректировки лучшей цены на портале
		public double Correct;

        // Не переоценивать китай
        public bool IgnoreChina;

        // Если запрещена продажа в минус, ниже этой границы переоценитьвать не дает
        public double LowerLimit;
	}
	public class Orders
	{
		// Значение параметра Supplier по умолчанию у ЗЧ
		public string DefaultSupplierName;

		// Категории хранения для заказа
		public List<string> StorageCategorysToOrder;

		// Список производителей для игнорирования при заказе
		public List<string> IgnoreToOrderList;

		// Список поставщиков для игнорирования при заказе
		public List<string> IgnoreSupplierToOrderList;
	}
}
