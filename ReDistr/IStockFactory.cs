using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReDistr
{
    public interface IStockFactory
    {
        Stock GetStock(string stockSignature);      //возвращает экземпляр Stock или null
                                                    //stockName дескриптор склада
        Stock TryGetStock(string inputString); //ищет в аргументе известные сигнатуры и создает соотвествующий экземпляр Stock для первого вхождения
        bool SetStockParams(string stockName, uint minimum, uint maximum, string signature, uint priority, List<string> mainManufacturers); //устанавливает параметры для создаваемых экземляров
                                                                                           //возвращает false если такая сигнатура уже существует
        IEnumerable<Stock> GetAllStocks();
	    void ClearStockParams();
    }
}
