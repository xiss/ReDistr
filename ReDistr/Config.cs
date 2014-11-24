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

        // Полный путь к текущей книге
        public static string PuthToThisWb;

		// Имя книги с параметрами
	    public static string NameOfParametersWb;

		// Количество складов
	    public static uint StockCount;

    }
}
