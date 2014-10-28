using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReDistr
{
    // Класс склад
    public class Stock
    {
        // Название склада
        public string Name;

        // Остаток  до выполнения перемещений
        public uint inStock;

        // Остаток после выполнения перемещений
        public uint inStockBefore;

        // Минимальный остаток, кратен inKit
        public uint minStock;

        // Период для расчета мин остатка по умолчанию
        public uint defaultPeriodMinStock;

        // Период для расчета макс остатка по умолчанию
        public uint defaultPeriodMaxStock;

        // Максимальный остаток, кратен inKit
        public uint maxStock;

        // Процент продаж
        public double sailPersent;

        // Резерв, количество товара в резерве
        public uint inReserve;

        // Приоритет, в спорных ситуациях используется для определения реципиента. 
        public uint priority;

        // Свободный остаток, который можно перемещать с данного склада, 
        // если свободный остаток отличен от 0, склад может быть донором. 
        // Не может быть меньше 0, если minStock отличен от нуля
        public uint freeStock;

        // Потребность данного склада в товаре, всегда положительна, 
        // если need отличен от 0, склад нуждается в товаре и может быть реципиентом. 
        public uint need;

        // Исключить данный склад из распределения
        // Выставляет свободный остатко равным реальному остатку за вычетом резервов
        private bool excludeFromMoovings;
    }
}
