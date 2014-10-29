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
        private uint inKit;

        // Количество товара в упаковке
        private uint inBundle;

        // Конструктор
        // TODO Конструктор вроде как не нужен, нужна пропертисы
        public Item(string id1c, string article, string storageCat, string name, string manufacturer, uint inKit, uint inBundle)
        {
            id1c = this.Id1C;
            article = this.Article;
            storageCat = this.StorageCategory;
            name = this.Name;
            manufacturer = this.Manufacturer;
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
