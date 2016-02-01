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
	public partial class OrderLists
	{
		private void Лист5_Startup(object sender, System.EventArgs e)
		{
		}

		private void Лист5_Shutdown(object sender, System.EventArgs e)
		{
		}

		#region Код, созданный конструктором VSTO

		/// <summary>
		/// Обязательный метод для поддержки конструктора - не изменяйте
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(Лист5_Startup);
			this.Shutdown += new System.EventHandler(Лист5_Shutdown);
		}

		#endregion
		// Строка с которой начинается заполненеие данными
		private const uint ArrayRowFirstFillNumber = 2;
		// Первая колонка с которой выводятся параметры складов
		private const uint ArrayColumnFirstFillNumber = 9;
		// Количество параметров для склада
		private const uint StockParametrsCount = 7;
		private const int ItemParametrsCount = 9;

		private const Excel.XlBordersIndex XlEdgeRight = Excel.XlBordersIndex.xlEdgeRight;
		private const Excel.XlBorderWeight XlThin = Excel.XlBorderWeight.xlThin;
		private const string NewNameStyle = "Хороший";
		private const string OldNameStyle = "Плохой";
		private const string NumberFormatText = "@";
		private const string NumberFormatPercentage = "0%";

		// Метод заполняет лист Test данными из словоря запчастей
		public void FillList(List<OrderRequiredItem> orderRequiredItems)
		{
			var curentRow = ArrayRowFirstFillNumber;
			var curentColumn = ArrayColumnFirstFillNumber;
			var stockCount = Config.StockCount;
			// 7 колонок под описание товара 9 колонок под параметры склада
			var resultRange = new dynamic[orderRequiredItems.Count + ArrayRowFirstFillNumber, ArrayColumnFirstFillNumber + Config.StockCount * StockParametrsCount];
			Range["A3:AD4000"].ClearContents();
			Range["A3:E4000"].NumberFormat = NumberFormatText;

			// Заполняем заголовки ЗЧ
			resultRange[0, 0] = "Id1C";
			resultRange[0, 1] = "Name";
			resultRange[0, 2] = "Article";
			resultRange[0, 3] = "Manufacturer";
			resultRange[0, 4] = "Supplier";
			resultRange[0, 5] = "StorageCat";
			resultRange[0, 6] = "inBundle";
			resultRange[0, 7] = "inKit";
			resultRange[0, 8] = "Note";

			// Выводим заголовки для параметров
			resultRange[0, curentColumn] = "Count";
			resultRange[0, curentColumn += stockCount] = "InReserve";
			resultRange[0, curentColumn += stockCount] = "SelCount";
			resultRange[0, curentColumn += stockCount] = "SailPers";
			resultRange[0, curentColumn += stockCount] = "Exclude";
			resultRange[0, curentColumn += stockCount] ="Old";
			resultRange[0, curentColumn += stockCount] = "Recomend";

			// Выводим заголовки складов
			curentColumn = ArrayColumnFirstFillNumber;
			for (var i = 0; i < StockParametrsCount; i++)
			{
				foreach (var stock in orderRequiredItems.First().Item.Stocks)
				{
					resultRange[1, curentColumn] = stock.Name.Substring(0, 1);
					curentColumn++;
				}
			}
			foreach (var order in orderRequiredItems)
			{
				// Выводим информацию по ЗЧ
				resultRange[curentRow, 0] = order.Item.Id1C;
				resultRange[curentRow, 1] = order.Item.Name;
				resultRange[curentRow, 2] = order.Item.Article;
				resultRange[curentRow, 3] = order.Item.Manufacturer;
				resultRange[curentRow, 4] = order.Item.Supplier;
				resultRange[curentRow, 5] = order.Item.StorageCategory;
				resultRange[curentRow, 6] = order.Item.InBundle;
				resultRange[curentRow, 7] = order.Item.InKit;
				resultRange[curentRow, 8] = "";


				// Выводим информацию по складам
				curentColumn = ArrayColumnFirstFillNumber;
				uint curentStock = 1;
				foreach (var stock in order.Item.Stocks)
				{
					resultRange[curentRow, curentColumn] = stock.Count;
					resultRange[curentRow, curentColumn += stockCount] = stock.InReserve;
					resultRange[curentRow, curentColumn += stockCount] = stock.CountSelings;
					resultRange[curentRow, curentColumn += stockCount] = stock.SailPersent;
					resultRange[curentRow, curentColumn += stockCount] = ReDistr.BoolToInt(stock.ExcludeFromMoovings);
					resultRange[curentRow, curentColumn += stockCount] = ReDistr.BoolToInt(order.Item.RequiredAvailability);
					// Если order содержит данный склад, выводим новый параметр RequiredAvailability
					if (order.OrderRequiredStocks.Contains(stock))
					{
						resultRange[curentRow, curentColumn += stockCount] = ReDistr.BoolToInt(true);
					}

					curentColumn = ArrayColumnFirstFillNumber;
					curentColumn += curentStock;
					curentStock++;
				}
				curentRow++;
			}

			// Выводим результат на лист
			Range[Cells[1, 1], Cells[orderRequiredItems.Count + 2, ArrayColumnFirstFillNumber + StockParametrsCount * Config.StockCount]].Value2 = resultRange;
			// Применяем стили и форматирование
			//Range[Cells[ArrayRowFirstFillNumber + 1, ItemParametrsCount + Config.StockCount * 5 + 1], Cells[orderRequiredItems.Count + 2, ItemParametrsCount + Config.StockCount * 6]].Style = OldNameStyle;
			//Range[Cells[ArrayRowFirstFillNumber + 1, ItemParametrsCount + Config.StockCount * 6 + 1], Cells[orderRequiredItems.Count + 2, ItemParametrsCount + Config.StockCount * 7]].Style = NewNameStyle;
			Range[Cells[ArrayRowFirstFillNumber + 1, ItemParametrsCount + Config.StockCount * 3 + 1], Cells[orderRequiredItems.Count + 2, ItemParametrsCount + Config.StockCount * 4]].NumberFormat = NumberFormatPercentage;
			// Границы колонок
			Range[Cells[ArrayRowFirstFillNumber + 1, ItemParametrsCount + 1], Cells[orderRequiredItems.Count + 2, ItemParametrsCount + Config.StockCount]].Borders(XlEdgeRight).Weight = XlThin;
			Range[Cells[ArrayRowFirstFillNumber + 1, ItemParametrsCount + 1], Cells[orderRequiredItems.Count + 2, ItemParametrsCount + Config.StockCount * 2]].Borders(XlEdgeRight).Weight = XlThin;
			Range[Cells[ArrayRowFirstFillNumber + 1, ItemParametrsCount + 1], Cells[orderRequiredItems.Count + 2, ItemParametrsCount + Config.StockCount * 3]].Borders(XlEdgeRight).Weight = XlThin;
			Range[Cells[ArrayRowFirstFillNumber + 1, ItemParametrsCount + 1], Cells[orderRequiredItems.Count + 2, ItemParametrsCount + Config.StockCount * 4]].Borders(XlEdgeRight).Weight = XlThin;
			Range[Cells[ArrayRowFirstFillNumber + 1, ItemParametrsCount + 1], Cells[orderRequiredItems.Count + 2, ItemParametrsCount + Config.StockCount * 5]].Borders(XlEdgeRight).Weight = XlThin;
			Range[Cells[ArrayRowFirstFillNumber + 1, ItemParametrsCount + 1], Cells[orderRequiredItems.Count + 2, ItemParametrsCount + Config.StockCount * 6]].Borders(XlEdgeRight).Weight = XlThin;
			Range[Cells[ArrayRowFirstFillNumber + 1, ItemParametrsCount + 1], Cells[orderRequiredItems.Count + 2, ItemParametrsCount + Config.StockCount * 7]].Borders(XlEdgeRight).Weight = XlThin;

			Select();
		}
	}
}
