using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace ReDistr
{
    public class SimpleStockFactory: IStockFactory
    {
        private static IStockFactory _factoryReference;

        private SimpleStockFactory()
        {
            _paramsList = new List<StockParams>();
        }

        public static IStockFactory CurrentFactory
        {
            get
            {
                if(_factoryReference == null)
                    _factoryReference = new SimpleStockFactory();
                return _factoryReference;
            }
        }
        public Stock GetStock(string stockSignature)
        {
            //TODO: проверить поведение при несущестующем значении
            try
            {
                var curentParams = _paramsList.Find(s => s.Signature == stockSignature);
                if(string.IsNullOrEmpty(curentParams.Signature)) return null;
                var stock = new Stock
                {
                    Signature =  curentParams.Signature,
                    Name = curentParams.Name,
                    defaultPeriodMaxStock = curentParams.Maximum,
                    defaultPeriodMinStock = curentParams.Minimum,
                    Priority = curentParams.Priority
                };
                return stock;
            }
            catch (Exception)
            {
                return null;
            }
            
        }



        public Stock TryGetStock(string inputString)
        {
            Stock foundStock = null;
            foreach (var prm in _paramsList)
            {
                if (inputString.ToLower().Contains(prm.Signature.ToLower()))
                {
                    foundStock = GetStock(prm.Signature);
                    break;
                }
            }
            return foundStock;
        }

        public bool SetStockParams(string stockName, uint minimum, uint maximum, string signature, uint priority)
        {
            if(_paramsList.Exists(s => s.Signature == signature)) return false;
            _paramsList.Add(new StockParams() { Maximum = maximum, Minimum = minimum, Name = stockName, Signature = signature, Priority = priority });
            return true;
        }

        public IEnumerable<Stock> GetAllStocks()
        {
            return _paramsList.Select(stockParam => GetStock(stockParam.Signature));
        }

        private readonly  List<StockParams> _paramsList;

        private struct StockParams
        {
            public string Name;
            public string Signature;
            public uint Maximum;
            public uint Minimum;
            public uint Priority;
        }
        
    }

    
}
