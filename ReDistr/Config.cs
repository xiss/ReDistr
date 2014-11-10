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
        public static DateTime periodSellingFrom;

        // Дата окончания периода продаж
        public static DateTime periodSellingTo;

        // Количество дней в периоде продаж
        public static int sellingPeriod;

        // Имя книги с остатками
        public static string NameOfStocksWB;

        // Имя книги с продажами
        public static string NameOfSealingsWB;

        // Полный путь к текущей книге
        public static string PuthToThisWB;

		// Имя книги с параметрами
	    public static string NameOfParametersWb;

		// Количество складов
	    public static uint StockCount;

    }
}
