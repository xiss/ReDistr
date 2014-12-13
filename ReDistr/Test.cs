using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ReDistr
{
	public partial class Test
	{
		private void Лист2_Startup(object sender, System.EventArgs e)
		{
		}

		private void Лист2_Shutdown(object sender, System.EventArgs e)
		{
		}

		#region Код, созданный конструктором VSTO

		/// <summary>
		/// Обязательный метод для поддержки конструктора - не изменяйте
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(Лист2_Startup);
			this.Shutdown += new System.EventHandler(Лист2_Shutdown);
		}

		#endregion

		private const int ArrayRowFirstFillNumber = 2;
		private const int ArrayColumnFirstFillNumber = 7;

		// Метод заполняет лист Test данными из словоря запчастей
		public void FillTestList(Dictionary<string, Item> items)
		{
			var curentRow = ArrayRowFirstFillNumber;
			var curentColumn = ArrayColumnFirstFillNumber;
			// 7 колонок под описание товара 10 колонок под склад
			var resultRange = new dynamic[items.Count + 2, 7 + Config.StockCount * 10];
			Cells.ClearContents();

			// Заполняем заголовки
			resultRange[0, 0] = "Id1C";
			resultRange[0, 1] = "Name";
			resultRange[0, 2] = "Article";
			resultRange[0, 3] = "Manufacturer";
			resultRange[0, 4] = "StorageCategory";
			resultRange[0, 5] = "inBundle";
			resultRange[0, 6] = "inKit";

			// Выводим заголовки для складов
			foreach (var stock in items.Values.First().Stocks)
			{
				resultRange[0, curentColumn] = stock.Name;
				resultRange[1, curentColumn] = "Count";
				resultRange[1, curentColumn += 1] = "InReserve";
				resultRange[1, curentColumn += 1] = "SelingsCount";
				resultRange[1, curentColumn += 1] = "SailPersent";
				resultRange[1, curentColumn += 1] = "MinStock";
				resultRange[1, curentColumn += 1] = "MaxStock";
				resultRange[1, curentColumn += 1] = "FreeStock";
				resultRange[1, curentColumn += 1] = "Need";
				resultRange[1, curentColumn += 1] = "Priority";
				resultRange[1, curentColumn += 1] = "Exclude";
				curentColumn++;
			}

			foreach (KeyValuePair<string, Item> item in items)
			{
				// Выводим информацию по ЗЧ
				resultRange[curentRow, 0] = item.Value.Id1C;
				resultRange[curentRow, 1] = item.Value.Name;
				resultRange[curentRow, 2] = item.Value.Article;
				resultRange[curentRow, 3] = item.Value.Manufacturer;
				resultRange[curentRow, 4] = item.Value.StorageCategory;
				resultRange[curentRow, 5] = item.Value.InBundle;
				resultRange[curentRow, 6] = item.Value.InKit;

				// Выводим информацию по складам
				curentColumn = ArrayColumnFirstFillNumber;
				foreach (var stock in item.Value.Stocks)
				{
					resultRange[curentRow, curentColumn] = stock.Count;
					resultRange[curentRow, curentColumn += 1] = stock.InReserve;
					resultRange[curentRow, curentColumn += 1] = stock.SelingsCount;
					resultRange[curentRow, curentColumn += 1] = stock.SailPersent;
					resultRange[curentRow, curentColumn += 1] = stock.MinStock;
					resultRange[curentRow, curentColumn += 1] = stock.MaxStock;
					resultRange[curentRow, curentColumn += 1] = stock.FreeStock;
					resultRange[curentRow, curentColumn += 1] = stock.Need;
					resultRange[curentRow, curentColumn += 1] = stock.Priority;
					resultRange[curentRow, curentColumn += 1] = stock.ExcludeFromMoovings;
					curentColumn++;
				}

				curentRow++;
			}

			// Выводим результат на лист
			Range[Cells[1, 1], Cells[items.Count + 2, 7 + 10 * Config.StockCount]].Value2 = resultRange;
		}
	}
}
