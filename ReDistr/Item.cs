using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReDistr
{
    // Класс запчасть
    public class Item
    {
        // Код 1С, уникален
        public string id1c;

        // Артикул
        public string article;

        // Категория хранения
        public string storageCat;

        // Название 
        public string name;

        // Производитель 
        public string manufacturer;

        // Количество товара в комплекте, не может быть равен 0, больше 0.
        public uint inKit;

        // Количество товара в упаковке
        public uint inBundle;

        // Конструктор
        public Item(string id1c, string article, string storageCat, string name, string manufacturer, uint inKit, uint inBundle)
        {
            id1c = this.id1c;
            article = this.article;
            storageCat = this.storageCat;
            name = this.name;
            manufacturer = this.manufacturer;
            inKit = this.inKit;
            inBundle = this.inBundle;

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
