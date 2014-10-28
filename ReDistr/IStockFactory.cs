using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReDistr
{
    interface IStockFactory
    {
        IStockFactory CurrentFactory { get; } //возвращает ссылку на экземпляр фабрики
        Stock GetStock(string stockSignature);     //возвращает экземпляр Stock или null
                                              //stockName дескриптор склада
        bool SetStockParams(string stockName, int minimum, int maximum, string signature); //устанавливает параметры для создаваемых экземляров
                                                                                           //возвращает false если такая сигнатура уже существует
    }
}
