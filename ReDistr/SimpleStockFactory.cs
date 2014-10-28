using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace ReDistr
{
    class SimpleStockFactory: IStockFactory
    {
        private IStockFactory _factoryReference;

        private SimpleStockFactory()
        {
            _paramsDictionary = new Dictionary<string, StockParams>();
        }

        public IStockFactory CurrentFactory
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
            try
            {
                var curentParams = _paramsDictionary[stockSignature];
                var stock = new Stock
                {
                    Name = curentParams.Name,
                    defaultPeriodMaxStock = curentParams.Maximum,
                    defaultPeriodMinStock = curentParams.Minimum
                };
                return stock;

            }
            catch (Exception)
            {
                return null;
            }
            
        }

        public bool SetStockParams(string stockName, uint minimum, uint maximum, string signature)
        {
            try
            {
                _paramsDictionary.Add(signature,
                    new StockParams() {Maximum = maximum, Minimum = minimum, Name = stockName});
                return true;
            }
            catch(ArgumentException)
            {
                return false;
            }
        }

        private readonly Dictionary<string, StockParams> _paramsDictionary;

        private struct StockParams
        {
            public string Name;
            public uint Maximum;
            public uint Minimum;
        }
        
    }

    
}
