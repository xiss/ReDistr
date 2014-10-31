using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ReDistr
{
    class Parser
    {
        // Указываем ячейки с настройками для парсера
        private const string RangeNameOfSealingsWb = "B14";
        private const string RangeNameOfStocksWb = "B13";
        private const string RangePuthToThisWb = "B15";
        private const uint RowNumberStockConfig = 4;
        private const uint RowNumberStocks = 7;

        // Получаем параметры с листа настроек
        // TODO надобы все методы кроме parse сделать приватными, но тогда ошибка в фабрике
        public Config GetConfig(Control control)
        {
            // Выбираем лист с настройками
            Globals.Control.Activate();
            
            // Создаем экземпляр класса Config для возврата из метода
            var config = new Config
            {
                NameOfSealingsWB = control.Range[RangeNameOfSealingsWb].Value2,
                NameOfStocksWB = control.Range[RangeNameOfStocksWb].Value2,
                PuthToThisWB = control.Range[RangePuthToThisWb].Value2
            };

            var curentRow = RowNumberStockConfig;
            do
            {
                // TODO Тут както надо заюзать фабрику
                curentRow++;
                // TODO Как проверить что ячейка пуста? Есть ли аналог isEmpty?
            } while (control.Range["B" + curentRow].Value.ToString() == "");

            return config;

            //control.Application.Workbooks.Open(control.Application.ActiveWorkbook.Path + "/" + control.Range["B9"].Value);
        }

       // Получаем остатки по складам
       public Item[] GetStocks(Control control, Config config)

       {
           // открываем нужную книгу
           control.Application.Workbooks.Open(config.PuthToThisWB + config.NameOfStocksWB);

           //Item = new Item();
           var curentRow = RowNumberStocks;

           do
           {

               curentRow++;
           } while (control.Range["B" + curentRow].Value.ToString() == "");

        
       }

        // Получаем данные по продажам
        public void GetSellings(Item item)
        {
            
        }

        // Получаем дополнительные параметры
        public void GetAdditionalParameters()
        {
            
        }

        // Основной метод парсера, из него вызываются все остальные
        // TODO что он должен возвращать? После парсинга получится и конфиг и массив итемов, как вернуть и то и другое?
        public void parse()
        {
            
        }

    }
}
