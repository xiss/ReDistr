using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
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

        // Свойство ЗЧ 2
        public string Property;

        // Свойство ЗЧ 1
        public string Property1;

        // Название 
        public string Name;

        // Производитель 
        public string Manufacturer;

        // Поставщик
        public string Supplier = Config.Config.Inst.Orders.DefaultSupplierName;

        // Количество товара в комплекте, не может быть равен 0, больше 0.
        public double InKit = 1;

        // Количество товара в упаковке
        public double InBundle = 1;

        // Коментарий по переоценке
        public string NoteReval = "";

        // Себестоимость
        // public double GetAVGCostPrice() = 0;

        /// <summary>
        /// Стоимость в ИМ
        /// </summary>
        public double Price = 0;

        // Колличество дней перезатарки сумарно по всем складам
        public double OverStockDaysForAllStocks;

        // Комментарий, почему установлена RequiredAvailability
        // public string NoteRequiredAvailability;

        // Остатки на складах
        public List<Stock> Stocks = new List<Stock>();

        // Конкуренты в Питерплюсе
        public List<Сompetitor> Сompetitors = new List<Сompetitor>();

        // Предустановленная цена
        public double PrePrice = 0;

        // Признак обязательного наличия ЗЧ на данном складе
        public bool RequiredAvailability;

        // В пути (заказано у поставщика)
        public double InTransfer;

        // Возвращает список всех возможных доноров, отсортированный по убыванию. Если задан список перемещений, то доноры выдаются из этого списка
        public List<Stock> GetListOfPossibleDonors(List<Transfer> existTransfers = null)
        {
            var listOfPossibleDonors = new List<Stock>();
            // Если список не задан, выдаем всех возможных доноров
            if (existTransfers == null)
            {
                // Если свободный осток отличен от нуля, то склад донор
                listOfPossibleDonors =
                    Stocks.Where(stock => stock.FreeStock > 0).OrderByDescending(stock => stock.FreeStock).ToList();
            }
            // Если список задан выдаем доноров из него
            else
            {
                foreach (var stock in Stocks)
                {
                    foreach (var transfer in existTransfers)
                    {
                        if (stock == transfer.StockFrom)
                        {
                            listOfPossibleDonors.Add(stock);
                        }
                    }
                }
            }

            // Если задана дериктива одного донора, то оставляем только его в списке
            if (Config.Config.OneDonor != null)
            {
                for (var i = 0; i < listOfPossibleDonors.Count; i++)
                {
                    if (listOfPossibleDonors[i] != Config.Config.OneDonor)
                    {
                        listOfPossibleDonors.Remove(listOfPossibleDonors[i]);
                        i--;
                    }
                }
            }
            return listOfPossibleDonors;
        }

        // Возвращает сумму всех свободных остатков, если задан список перемещений то остатки берутся из доноров в этих перемещениях
        public double GetSumFreeStocks(List<Transfer> existTransfers = null)
        {
            // Если задан OneDonor, выдаем свободные остатки только для этого донора
            if (Config.Config.OneDonor != null)
            {
                return Stocks.Where(stock => stock == Config.Config.OneDonor).Sum(stock => stock.FreeStock);
            }

            // Если список не задан, выдаем сумму для всех складов
            if (existTransfers == null)
            {
                return Stocks.Sum(stock => stock.FreeStock);
            }

            // Если задан список доноров, выдаем сумму свободных остатков по этим донорам
            var existDonors = new List<Stock>();
            foreach (var stock in Stocks)
            {
                foreach (var transfer in existTransfers)
                {
                    if (stock == transfer.StockFrom)
                    {
                        existDonors.Add(stock);
                    }
                }
            }

            return existDonors.Sum(stock => stock.FreeStock);
        }

        // Возвращает общее количество ЗЧ без учета резервов
        public double GetSumStocks(bool withReserve = true, bool withInTransfer = false)
        {
            double sumStocks;

            // Если нужно учитываем резервы
            if (withReserve)
            {
                sumStocks = Stocks.Sum(stock => stock.Count - stock.InReserve);
            }
            else
            {
                sumStocks = Stocks.Sum(stock => stock.Count);
            }

            // Если нужно добавляем количество в пути
            if (withInTransfer)
            {
                sumStocks = sumStocks + InTransfer;
            }

            if (sumStocks < 0)
            {
                sumStocks = 0;
            }
            return sumStocks;
        }

        // Возвращает среднюю себестоимость
        public double GetAVGCostPrice()
        {
            var a = Stocks.Sum(stock => stock.CostPrice);
            var b = GetSumStocks(false);
            var c = Math.Round(Stocks.Sum(stock => stock.CostPrice * stock.Count) / GetSumStocks(false), 2);
            return Math.Round(Stocks.Sum(stock => stock.CostPrice * stock.Count) / GetSumStocks(false), 2);
        }

        // Возвращает общий минимальный остаток
        public double GetSumMinStocks()
        {
            var sumMinStocks = Stocks.Sum(stock => stock.MinStock);
            return sumMinStocks;
        }

        // Возвращает общий максимальный остаток
        public double GetSumMaxStocks()
        {
            var sumMaxStocks = Stocks.Sum(stock => stock.MaxStock);
            return sumMaxStocks;
        }

        // Возвращает сумму продаж
        public double GetSumSelings(bool inKits = false)
        {
            var sumSelings = Stocks.Where(stock => stock.CountSelings > 0).Sum(stock => stock.CountSelings);

            // Переводим в комплекты
            if (inKits)
            {
                sumSelings = sumSelings / InKit;
            }

            return sumSelings;
        }

        // Обновляет свободные остатки
        public void UpdateFreeStocks(string typeFreeStock)
        {
            foreach (var stock in Stocks)
            {
                stock.UpdateFreeStock(this, typeFreeStock);
            }
        }

        // Возвращает ближаещего конкурента с учетом исключений
        public Сompetitor GetСompetitor(bool withDeliveryTime, bool withCompetitorsStocks, bool withExcludes = true, bool checkDumping = false)
       {
            var sumStocks = GetSumStocks();
            var dumpingPersent = Config.Config.Inst.Revaluations.DumpingPersent;
            var maxCompetitorsToMiss = Config.Config.Inst.Revaluations.MaxCompetitorsToMiss;
            double deliveryTime = Config.Config.Inst.Revaluations.OurDeliveryTime + Config.Config.Inst.Revaluations.DeltaDeliveryTime;


            Сompetitors = Сompetitors.OrderBy(competitor => competitor.PositionNumber).ToList();

            var i = 0;

            for (int k = 0; k < Сompetitors.Count; k++)
            {
                // Проверяем список исключений если конкуреты из этого списка переходим к следующему
                if (Config.Config.Inst.Revaluations.ListExcludeCompetitors.Contains(Сompetitors[k].Id) & withExcludes)
                {
                    NoteReval = NoteReval + k + " В списке исключений " + Сompetitors[k].Id + "\n";
                    continue;
                }
                i++;

                // Проверяем на демпинг
                // Проверяем есть ли следующий конурент
                if (Сompetitors.Exists(competitor => competitor.PositionNumber == Сompetitors[k].PositionNumber + 1))
                {
                    //только первого конкурента
                    if (checkDumping & i == 1 & Сompetitors[k].PriceWithAdd * (1 + dumpingPersent) < Сompetitors[k + 1].PriceWithAdd & maxCompetitorsToMiss != 0 & maxCompetitorsToMiss >= i)
                    {
                        NoteReval = NoteReval + k + " Демпинг " + Сompetitors[k].Id + "\n";
                        continue;
                    }
                }

                // Проверяем срок поставки, если не соответствует переходим к следующему
                if (Сompetitors[k].DeliveryTime > deliveryTime & withDeliveryTime & maxCompetitorsToMiss != 0 & maxCompetitorsToMiss >= i)
                {
                    NoteReval = NoteReval + k + " Большой срок поставки " + Сompetitors[k].Id + "\n";
                    continue;
                }

                // Проверяем запас, если он меньше необходимого переходим к следующему
                if (Сompetitors[k].Count < sumStocks * Config.Config.Inst.Revaluations.DeltaCompetitorStock & withCompetitorsStocks & maxCompetitorsToMiss != 0 & maxCompetitorsToMiss >= i)
                {
                    NoteReval = NoteReval + k + " Остаток " + Сompetitors[k].Id + "\n";
                    continue;
                }
                // Проверяем чтобы регион не содержал слово уценка
                if (Сompetitors[k].Region.Contains("Уценка"))
                {
                    NoteReval = NoteReval + k + " Регион содержит слово (Уценка) " + Сompetitors[k].Id + "\n";
                    continue;
                }
                // Проверяем чтобы он не был первым
                if (Сompetitors[k].PositionNumber == 1)
                {
                    //continue;
                }
                return Сompetitors[k];
            }
            return null;
        }

        ///<summary>Возвращает наш срок поставки</summary>
        public double GetOurDeliveryTime()
        {
            double ourDeliveryTime = 0;
            if (Сompetitors.Exists(competitor => competitor.Id == "Наш прайс"))
            {
                ourDeliveryTime = Сompetitors.Find(competitor => competitor.Id == "Наш прайс").DeliveryTime;
            }
            return ourDeliveryTime;
        }

        ///<summary>Возвращает нашу цену на портале с наценкой</summary>
        public double GetPricePortalWithAdd()
        {
            double pricePortalWithAdd = 0;
            if (Сompetitors.Exists(competitor => competitor.Id == "Наш прайс"))
            {
                pricePortalWithAdd = Сompetitors.Find(competitor => competitor.Id == "Наш прайс").PriceWithAdd;
            }
            return pricePortalWithAdd;
        }
        // Возвращает новую цену расчитанную опираясь на указанного конкурента
        public double GetNewPrice(Сompetitor сompetitor, bool allowSellingLoss)
        {
            var correct = Config.Config.Inst.Revaluations.Correct;
            double markup = 1;
            // Расчитываем новую цену
            double newPrice = 0;

            // Если есть предустановленная цена, используем ее
            if (PrePrice != 0)
            {
                NoteReval = NoteReval + "\n Предустановленная цена (" + PrePrice + ")";
                return PrePrice;
            }

            // Если конкурент есть
            if (сompetitor != null)
            {
                //Не Китай
                if (Manufacturer != "Китай")
                {
                    switch (StorageCategory)
                    {
                        case "Везде":
                        case "Нигде":
                        case "МинЗапас":
                            switch (Manufacturer)
                            {
                                case "GREAT WALL":
                                case "CHERY":
                                case "LIFAN":
                                case "GEELY":
                                case "ZOTYE":
                                case "CASTROL":
                                    markup = 1.2;
                                    break;
                                case "ZEKKERT":
                                    markup = 1.15;
                                    break;
                                default:
                                    markup = 1.4;
                                    break;
                            }

                            if (сompetitor.PriceWithoutAdd < GetAVGCostPrice() * markup)
                            {
                                newPrice = GetAVGCostPrice() * markup;
                            }
                            else
                            {
                                newPrice = сompetitor.PriceWithoutAdd * correct;
                            }
                            NoteReval = NoteReval + "\n Наценка (" + markup + ")";
                            break;
                        case "НЛ12":
                            if (сompetitor.PriceWithoutAdd > GetAVGCostPrice() * 0.75)
                            {
                                newPrice = GetAVGCostPrice() * 0.75;
                            }
                            else
                            {
                                newPrice = сompetitor.PriceWithoutAdd * correct;
                            }
                            NoteReval = NoteReval + "\n Наценка (НЛ12) (" + 0.75 + ")";
                            break;
                        case "НЛ24":
                            if (сompetitor.PriceWithoutAdd > GetAVGCostPrice() * 0.5)
                            {
                                newPrice = GetAVGCostPrice() * 0.5;
                            }
                            else
                            {
                                newPrice = сompetitor.PriceWithoutAdd * correct;
                            }
                            NoteReval = NoteReval + "\n Наценка (НЛ24) (" + 0.5 + ")";
                            break;
                        case "НЛ6":
                            if (сompetitor.PriceWithoutAdd < GetAVGCostPrice())
                            {
                                newPrice = GetAVGCostPrice();
                            }
                            else
                            {
                                newPrice = сompetitor.PriceWithoutAdd * correct;
                            }
                            NoteReval = NoteReval + "\n Наценка (Попова) (" + 1 + ")";
                            break;
                        default:
                            newPrice = сompetitor.PriceWithoutAdd * correct;
                            break;
                    }
                }
                //Китай
                else
                {
                    switch (Property)
                    {
                        case "НЛ12":
                        case "НЛ6":
                        case "НЛ24":
                            if (newPrice > GetAVGCostPrice() * 1.2)
                            {
                                newPrice = GetAVGCostPrice() * 1.2;
                            }
                            NoteReval = NoteReval + "\n Наценка (НЛ) не больше (" + 1.2 + ")";
                            break;
                        default:
                            newPrice = сompetitor.PriceWithoutAdd * correct;
                            if (newPrice < GetAVGCostPrice() * 1.05)
                            {
                                //newPrice = GetAVGCostPrice() * 1.05;
                            }
                            if (newPrice > GetAVGCostPrice() * 2)
                            {
                                //newPrice = GetAVGCostPrice() * 2;
                            }
                            break;
                    }
                }
            }
            // Если конкурента нет
            else
            {
                // Китай
                if (Manufacturer == "Китай")
                {
                    // Вариант с лесницей по себестоимости
                    if (GetAVGCostPrice() > 0 & GetAVGCostPrice() < 80)
                    {
                        newPrice = GetAVGCostPrice() * 4;
                    }
                    else if (GetAVGCostPrice() > 80 & GetAVGCostPrice() < 200)
                    {
                        newPrice = GetAVGCostPrice() * 3;
                    }
                    else if (GetAVGCostPrice() > 201 & GetAVGCostPrice() < 500)
                    {
                        newPrice = GetAVGCostPrice() * 2;
                    }
                    else if (GetAVGCostPrice() > 501 & GetAVGCostPrice() < 1000)
                    {
                        newPrice = GetAVGCostPrice() * 1.9;
                    }
                    else if (GetAVGCostPrice() > 1001 & GetAVGCostPrice() < 2000)
                    {
                        newPrice = GetAVGCostPrice() * 1.7;
                    }
                    else if (GetAVGCostPrice() > 2001 & GetAVGCostPrice() < 4000)
                    {
                        newPrice = GetAVGCostPrice() * 1.5;
                    }
                    else if (GetAVGCostPrice() > 4001 & GetAVGCostPrice() < 8000)
                    {
                        newPrice = GetAVGCostPrice() * 1.2;
                    }
                    else if (GetAVGCostPrice() > 8001 & GetAVGCostPrice() < 15000)
                    {
                        newPrice = GetAVGCostPrice() * 1.15;
                    }
                    else if (GetAVGCostPrice() > 15001 & GetAVGCostPrice() < 1000000000)
                    {
                        newPrice = GetAVGCostPrice() * 1.1;
                    }
                }
                // Не китай
                else
                {
                    switch (StorageCategory)
                    {
                        case "Везде":
                        case "Нигде":
                        case "МинЗапас":
                            switch (Manufacturer)
                            {
                                case "GREAT WALL":
                                case "CHERY":
                                case "LIFAN":
                                case "GEELY":
                                case "ZOTYE":
                                case "CASTROL":
                                    markup = 1.2;
                                    break;
                                case "ZEKKERT":
                                    markup = 1.15;
                                    break;
                                default:
                                    markup = 1.4;
                                    break;
                            }
                            newPrice = GetAVGCostPrice() * markup;
                            break;
                        case "НЛ12":
                            newPrice = GetAVGCostPrice() * 0.85;
                            break;
                        case "НЛ24":
                            newPrice = GetAVGCostPrice() * 0.6;
                            break;
                        case "НЛ6":
                            newPrice = GetAVGCostPrice();
                            break;
                        default:
                            newPrice = GetAVGCostPrice() * 1.4;
                            break;
                    }
                }
            }
            // Если новая цена ниже себестоимости, возвращаем себестоимость
            if (newPrice < (GetAVGCostPrice() * Config.Config.Inst.Revaluations.LowerLimit) && !allowSellingLoss)
            {
                newPrice = GetAVGCostPrice() * Config.Config.Inst.Revaluations.LowerLimit;
            }
            return Math.Round(newPrice, 2);
        }
    }
}