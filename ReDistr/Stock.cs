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
        public uint InStock;

        // Остаток после выполнения перемещений
        public uint InStockBefore;

        // Минимальный остаток, кратен inKit
        public uint MinStock;

        // Период для расчета мин остатка по умолчанию
        public uint defaultPeriodMinStock;

        // Период для расчета макс остатка по умолчанию
        public uint defaultPeriodMaxStock;

        // Максимальный остаток, кратен inKit
        public uint MaxStock;

        // Процент продаж
        private double sailPersent;

        // Резерв, количество товара в резерве
        private uint inReserve;

        // Приоритет, в спорных ситуациях используется для определения реципиента. 
        public uint Priority;

        // Свободный остаток, который можно перемещать с данного склада, 
        // если свободный остаток отличен от 0, склад может быть донором. 
        // Не может быть меньше 0, если minStock отличен от нуля
        public uint FreeStock;

        // Потребность данного склада в товаре, всегда положительна, 
        // если need отличен от 0, склад нуждается в товаре и может быть реципиентом. 
        public uint Need;

        // Исключить данный склад из распределения
        // Выставляет свободный остатко равным реальному остатку за вычетом резервов
        private bool excludeFromMoovings;
    }
}
