using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReDistr
{
    class Config
    {
        // Дата снятия отчета с остатками
        public DateTime StockDate;

        // Дата начала периода продаж
        public DateTime periodSellingFrom;

        // Дата окончания периода продаж
        public DateTime periodSellingTo;

        // Количество дней в периоде продаж
        public uint sellingPeriod;

        // Имя книги с остатками
        public string NameOfStocksWB;

        // Имя книги с продажами
        public string NameOfSealingsWB;

        // Полный путь к текущей книге
        public string PuthToThisWB;

    }
}
