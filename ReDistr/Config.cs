using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.IO;

namespace ReDistr
{
	[Serializable]
	public class Config
	{
		public FilesCfg FilesCfg = FilesCfg.Inst;
		public RevaluationsCfg RevaluationsCfg = RevaluationsCfg.Inst;
		public OrdersCfg OrdersCfg = OrdersCfg.Inst;
		public TransfersCfg TransfersCfg = TransfersCfg.Inst;

		/// <summary>
		/// Загрузить настройки 
		/// </summary>
		static Config()
		{
			try
			{
				var serializer = new XmlSerializer(typeof(Config)); using (var stream = File.OpenRead("config.xml"))
				{
					_inst = (Config)serializer.Deserialize(stream);
				}
			}
			catch (Exception e)
			{
				//TODO логирование
				//LogManager.GetCurrentClassLogger().Error("Ошибка загрузки настроек. {0}", e.Message);
				//TODO как закончить выполнение функции, вызваться может где угодно
			}
		}
		/// <summary>
		/// Сохранить настройки
		/// </summary>
		public static void Save()
		{
			try
			{
				var serializer = new XmlSerializer(typeof(Config));
				var i = new Config();
				Stream writer = new FileStream("config.xml", FileMode.OpenOrCreate);
				serializer.Serialize(writer, i);
				writer.Close();
			}
			catch (Exception e)
			{
				//LogManager.GetCurrentClassLogger().Error("Ошибка сохранения настроек. {0}", e.Message);
			}
		}

		//Singleton
		private Config() { }
		private static Config _inst;
		public static Config Inst => _inst ?? (_inst = new Config());

		// Дата снятия отчета с остатками
		public static DateTime StockDate;

		// Склад для перемещения выбранных категорий (неликвид)
		public static Stock StockToTransferSelectedStorageCategory = SimpleStockFactory.CurrentFactory.GetStock(TransfersCfg.Inst.StockNameToTransferSelectedStorageCategory);
		
		// Если параметр указан, то перемещения делать только с этого склада
		public static Stock OneDonor = SimpleStockFactory.CurrentFactory.GetStock(TransfersCfg.Inst.StockNameOneDonor);

		// Склад для оптовых отгрузок
		public static Stock WholesaleStock = SimpleStockFactory.CurrentFactory.GetStock(RevaluationsCfg.Inst.StockNameWholesaleStock);

		// Дата начала периода продаж
		public static DateTime PeriodSellingFrom;

		// Дата окончания периода продаж
		public static DateTime PeriodSellingTo;

		// Количество дней в периоде продаж
		public static int SellingPeriod;

		// Количество складов
		public static uint StockCount;

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

	[Serializable]
	public class FilesCfg
	{
		// Папка с перемещениями
		public  string FolderTransfers ;

		// Папка с архивом перемещений
		public  string FolderArchiveTransfers ;

		// Папка с переоценками
		public  string FolderRevaluations ;

		// Папка с архивом переоценками
		public  string FolderArchiveRevaluations ;

		// Имя книги с остатками
		public  string NameOfStocksWb;

		// Имя книги с продажами
		public  string NameOfSealingsWb;

		// Имя книги с конкурентами
		public  string NameOfCompetitorsWb;

		// Имя книги с параметрами
		public  string NameOfParametersWb;

		//Singleton
		private FilesCfg() { }
		private static FilesCfg _inst;
		public static FilesCfg Inst => _inst ?? (_inst = new FilesCfg());
	}
	[Serializable]

	public class OrdersCfg
	{
		// Значение параметра Supplier по умолчанию у ЗЧ
		public string DefaultSupplierName = "none";

		//Singleton
		private OrdersCfg() { }
		private static OrdersCfg _inst;
		public static OrdersCfg Inst => _inst ?? (_inst = new OrdersCfg());
	}

	[Serializable]

	public class TransfersCfg
	{
		// Категория обязательного наличия
		public  string NameOfStorageCatRequiredAvailability = "МинЗапас;Везде";

		// имя склада для перемещения выбранных категорий (неликвид)
		public string StockNameToTransferSelectedStorageCategory;

		// Список категорий для перемещения на выбранный склад
		public  List<String> ListSelectedStorageCategoryToTransfer;

		// Если параметр указан, то перемещения делать только с этого склада
		public  string StockNameOneDonor = null;

		// Категории хранения товара для перемещения
		public  List<string> ListStorageCategoryToTransfers;

		// Свойства ЗЧ для обязательного наличия, синоним мин. остатка (свойство ЗЧ1)
		public  List<string> ListPropertyRequiredAvailability;

		//Singleton
		private TransfersCfg() { }
		private static TransfersCfg _inst;
		public static TransfersCfg Inst => _inst ?? (_inst = new TransfersCfg());
	}
	[Serializable]

	public class RevaluationsCfg
	{
		// id нашего прайса на П+
		public  string IdPriceAp;

		// Склад для оптовых отгрузок
		public  string StockNameWholesaleStock;

		// Список конкурентов исключений
		public  List<string> ListExcludeCompetitors;

		// Процент остатка от нашего склада для рассмотрения его как конкурента
		public  double MinStockForCompetitor;

		// Наш срок поставки
		public  double OurDeliveryTime;

		// Процент для выявления демпинга
		public  double DumpingPersent;

		// Допустимая разница в сроке доставки между нами и конкурентом
		public  double DeltaDeliveryTime;

		// Отношение остатков конкурента к нашему
		public  double DeltaCompetitorStock;

		// Максимально можно пропустить конкурентов
		public  double MaxCompetitorsToMiss;

		// типа конкурента
		public  int TypeCompetitor;

		//Singleton
		private RevaluationsCfg() { }
		private static RevaluationsCfg _inst;
		public static RevaluationsCfg Inst => _inst ?? (_inst = new RevaluationsCfg());
	}
}
