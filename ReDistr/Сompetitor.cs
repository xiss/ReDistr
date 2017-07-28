namespace ReDistr
{
    // Конкуренты на порталах
    public class Сompetitor
    {
        // ЗЧ

        // Срок поставки
        public double Count;
        public double DeliveryTime;

        // Цена

        // Код поставщика
        public string Id;
        public Item Item;

        // Номер строки на портале
        public double PositionNumber;

        // Регион
        public string Region;
        private double _price;

        public double PriceWithoutAdd
        {
            get
            {
                double price = 0;
                // double ratio;
                // Вариант с пристрелкой
                // Вычисляем наценку конкурента
                // Проверяем есть ли у нас данные по прошлой цене в П+, если данных нет, используем шаблон
                //if (Item.Сompetitors.Exists(competitor => competitor.Id == "Наш прайс"))
                //{
                //    // Определяем реальную наценку конкурента
                //    double ourPrice = Item.Сompetitors.Find(competitor => competitor.Id == "Наш прайс")._price;
                //    ratio = ourPrice / Item.Price;
                //    // Проверка на максимум
                //    if (ratio > 1.14)
                //    {
                //        ratio = 1.14;
                //    }
                //    // Проверка на минимум
                //    if (ratio < 1.11)
                //    {
                //        ratio = 1.11;
                //    }
                //}
                //else
                //{
                //    ratio = 1.13;
                //}
                //price = _price / ratio;

                // Вариант с порогами
                switch (Config.TypeCompetitor)
                {
                    // Автопитер
                    case 1:
                        if (_price > 0 & _price <= 1146.49)
                        {
                            price = _price/1.1581;
                        }
                        else if (_price >= 1146.50 & _price <= 2244.05)
                        {
                            price = _price/1.1508;
                        }
                        else if (_price >= 2244.06 & _price <= 4517.00)
                        {
                            price = _price/1.1435;
                        }
                        else if (_price >= 4517.01 & _price <= 5624.76)
                        {
                            price = _price/1.1363;
                        }
                        else if (_price >= 5624.77 & _price <= 6739.62)
                        {
                            price = _price/1.1327;
                        }
                        else if (_price >= 6739.63 & _price <= 7847.30)
                        {
                            price = _price/1.1291;
                        }
                        else if (_price >= 7847.31 & _price <= 8947.83)
                        {
                            price = _price/1.1255;
                        }
                        else if (_price >= 8947.84 & _price <= 16828.86)
                        {
                            price = _price/1.1219;
                        }
                        else if (_price >= 16663 & _price <= 100000000)
                        {
                            price = _price/1.104;
                        }
                        break;
                    // Иксора
                    case 2:
                        if (_price > 0 & _price <= 220)
                        {
                            price = _price / 2.2;
                        }
                        else if (_price >= 220.01 & _price <= 5445)
                        {
                            price = _price / 1.089;
                        }
                        else if (_price >= 5445.01 & _price <= 7364.56)
                        {
                            price = _price / 1.078;
                        }
                        else if (_price >= 7364.57 & _price <= 100000000)
                        {
                            price = _price / 1.067;
                        }
                        break;
                    // Партком
                    case 3:
                        if (_price > 0 & _price <= 150)
                        {
                            price = _price / 2;
                        }
                        else if (_price >= 150.01 & _price <= 5634)
                        {
                            price = _price / 1.140;
                        }
                        else if (_price >= 5634.01 & _price <= 100000000)
                        {
                            price = _price / 1.1;
                        }
                        break;
                }

                return price;
            }
            set { _price = value; }
        }

        public double PriceWithAdd
        {
            get { return _price; }
        }

        // Количество
    }
}