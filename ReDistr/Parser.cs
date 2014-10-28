using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ReDistr
{
    class Parser
    {
        // Получаем параметры с листа настроек
        public void GetConfig(Control control)
        {
            control.Application.Workbooks.Open(control.Application.ActiveWorkbook.Path + "/" + control.Range["B9"].Value);
        }

        // Получаем остатки по складам
        public Item[] GetStocks(Control control)
        {
            
        }

        // Получаем данные по продажам
        public void GetSellings(Item item,)
        {
            
        }

        // Получаем дополнительные параметры
        public void GetAdditionalParameters()
        {
            
        }
    }
}
