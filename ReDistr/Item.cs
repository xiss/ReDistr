using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace ReDistr
{
    // Класс запчасть
    public class Item
    {
        // Код 1С, уникален
        public string Id1C;

        // Артикул
        public string Article;

        // Категория хранения
        public string StorageCategory;

        // Название 
        public string Name;

        // Производитель 
        public string Manufacturer;

        // Количество товара в комплекте, не может быть равен 0, больше 0.
        public double inKit = 1;

        // Количество товара в упаковке
        public double inBundle = 1;

        // Остатки на складах 
        public List<Stock> Stocks = new List<Stock>();

        // Проверяет существование такого склада, и если находит возвращает ссылку
        public Stock getStock()
        {
            return null;
        }

        // индексатор
        //public Stock this[string index]
        //{
        //    get
        //    {
        //        foreach (Stock st in Stocks)
        //        {
        //            if (st.name == index)
        //            {
        //                return st;
        //            }

        //        }
        //        return null;
        //    }
        //}

        //public Stock[] Stocks;
    }
}
