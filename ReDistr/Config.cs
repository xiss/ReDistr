using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReDistr
{
    class Config
    {
        // Массив с конфигурацией складов
        // 0 приоритет
        // 1 мин. ост
        // 2 макс ост
        //int[,] stockInt = new int[3, 3];

        public Tuple<string, int, int, string > StocksTuple;




        // Сигнатура массива
        //string[,] stockSign = new string[2,2];

        // Дата снятия отчета с остатками
        private DateTime StockDate;

        // Дата начала периода продаж
        private DateTime periodSellingFrom;

        // Дата окончания периода продаж
        private DateTime periodSellingTo;

        // Количество дней в периоде продаж
        private uint sellingPeriod;
    }
}
