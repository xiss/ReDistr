using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ReDistr
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        // Обновить блок с информацией
        public void updateInfo(){
            labelPeriodSelling.Label = Config.PeriodSellingFrom.ToString("dd.MM.yy") + " - " + Config.PeriodSellingTo.ToString("dd.MM.yy");
            labelPeriodSellingCount.Label = Config.SellingPeriod.ToString() + " (дни)";
            labelStockDate.Label = Config.StockDate.ToString("dd.MM.yy");
        }
        
        // Сформировать списки для заказа
        private void buttonGetOrderLists_Click(object sender, RibbonControlEventArgs e)
        {
            // Парсим данные из файлов
            var parser = new Parser();
            var items = parser.Parse();

            // Если парсинг не удался, выходим
            if (items == null)
            {
                return;
            }

            // Подготавливаем данные
            ReDistr.PrepareData(items);

            // Выводим таблицу для тестов
            Globals.Test.FillListStocks(items);

            // Обновляем параметры
            this.updateInfo();

            // Составляем список для заказа
            var orderRequiredItems = ReDistr.GetOrderRequiredItems(items);

            // Выводим списки для заказа на лист со списками
            Globals.OrderLists.FillList(orderRequiredItems);
        }
        
        // Сформировать заказы
        private void buttonGetOrders_Click(object sender, RibbonControlEventArgs e)
        {
            // Парсим данные из файлов
            var parser = new Parser();
            var items = parser.Parse();

            // Если парсинг не удался, выходим
            if (items == null)
            {
                return;
            }

            // Подготавливаем данные
            ReDistr.PrepareData(items);

            // Выводим таблицу для тестов
            Globals.Test.FillListStocks(items);

            // Обновляем параметры
            this.updateInfo();

            // Формирует заказы
            var orders = new List<Order>();
            orders = ReDistr.GetOrders(orders, items);

            // Выводим заказы на страницу заказов
            Globals.Orders.FillList(orders);

            // Выбираем лист с pfrfpfvb
            Globals.Orders.Select();
        }
        // Архивирует старый книги с перемещениями, и создает новые
        private void buttonMakeTransfersBook_Click(object sender, RibbonControlEventArgs e)
        {
            // Архивируем предыдущие перемещения
            ReDistr.ArchiveTransfers();

            // Создаем книги для импорта в Excel
            Globals.Transfers.MakeImportTransfers();
        }        
   
        // Сформировать перемещения
        private void buttonGetTransfers_Click(object sender, RibbonControlEventArgs e)
        {
            // Парсим данные из файлов
            var parser = new Parser();
            var items = parser.Parse();

            // Если парсинг не удался, выходим
            if (items == null)
            {
                return;
            }

            // Подготавливаем данные
            ReDistr.PrepareData(items);

            // Выводим таблицу для тестов
            Globals.Test.FillListStocks(items);

            // Формируем перемещения
            var transfers = new List<Transfer>();
            // для обеспечения одного комплекта
            transfers = ReDistr.GetTransfersFirstLvl(items, transfers);
            // для обеспечения мин. остатка
            transfers = ReDistr.GetTransfersSecondLvl(items, transfers);
            // для обеспечения необходимого запаса, перемещения создаются для уже созданных направлений
            transfers = ReDistr.GetTransfersThirdLvl(items, transfers);

            // Если необходимо делаем перемещение неликвида на Попова
            if (Config.StockToTransferSelectedStorageCategory != null)
            {
                transfers = ReDistr.GetTransfersIlliuid(items, transfers);
            }

            // Выводим отчет для тестов если необходимо
            if (Config.ShowReport)
            {
                Globals.Test.FillListTransfers(transfers, items);
            }

            // Выводим перемещения на лист для перемещений
            Globals.Transfers.FillList(transfers);

            // Обновляем параметры
            this.updateInfo();

            // Выбираем лист с перемещениями
            Globals.Transfers.Select();
        }

    }
}
