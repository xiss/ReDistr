using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ReDistr
{
	public partial class Orders
	{
		private void Лист4_Startup(object sender, System.EventArgs e)
		{
		}

		private void Лист4_Shutdown(object sender, System.EventArgs e)
		{
		}

		#region Код, созданный конструктором VSTO

		/// <summary>
		/// Обязательный метод для поддержки конструктора - не изменяйте
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(Лист4_Startup);
			this.Shutdown += new System.EventHandler(Лист4_Shutdown);
		}

		#endregion
		// Определяем настройки
		private const int StartRow = 3;
		private const int ItemParametrsCount = 8;
		private const Excel.XlBordersIndex XlEdgeRight = Excel.XlBordersIndex.xlEdgeRight;
		private const Excel.XlBorderWeight XlThin = Excel.XlBorderWeight.xlThin;
		private const string CountNameStyle = "Хороший";
		private const string HeaderNameStyle = "Заголовок 1";
		private const string NumberFormatText = "@";

		// Выводит на лист заказы сгруппированные по поставщику
		public void FillList(List<Order> orders)
		{
			// Очищаем лист
			// TODO /5 Подумать как сделать это проще, сейчас становятся активными лишние ячейки
			//Cells.Clear();
			Range["A3:Z1500"].Clear();
			//Range["A3:E1500"].NumberFormat = NumberFormatText;

			var curentRow = StartRow;
			var firstIteration = true;
			var resultRangeHeader = new dynamic[1, ItemParametrsCount + Config.StockCount * 4];

			// Перебираем список поставщиков
			foreach (var supplier in Config.ListSuppliers)
			{
				// Выбираем ЗЧ с данным поставщиком
				var supplierOrder = orders.Where(order => order.Item.Supplier == supplier).ToList();
				var resultRange = new dynamic[supplierOrder.Count + 1, ItemParametrsCount + Config.StockCount * 4];
				resultRange[0, 0] = supplier;
				var i = 1;
				// Добавляем ЗЧ в массив 
				foreach (var order in supplierOrder)
				{
					resultRange[i, 0] = order.Item.Id1C;
					resultRange[i, 1] = order.Item.Article;
					resultRange[i, 2] = order.Item.Name;
					resultRange[i, 3] = order.Item.Manufacturer;
					resultRange[i, 4] = order.Item.StorageCategory;
					resultRange[i, 5] = order.Item.InBundle;
					resultRange[i, 6] = order.Item.InKit;
					resultRange[i, 7] = order.Count;

					// Добавляем информацию по складам
					var y = ItemParametrsCount;
					foreach (var stock in order.Item.Stocks)
					{
						// Выводим заголовек склада
						if (firstIteration)
						{
							var shortName = stock.Name.Substring(0, 1);
							resultRangeHeader[0, y] = shortName;
							resultRangeHeader[0, y + Config.StockCount] = shortName;
							resultRangeHeader[0, y + Config.StockCount * 2] = shortName;
							resultRangeHeader[0, y + Config.StockCount * 3] = shortName;
						}
						resultRange[i, y] = stock.CountOrigin;
						resultRange[i, y + Config.StockCount] = stock.CountSelings;
						resultRange[i, y + Config.StockCount * 2] = stock.MinStock;
						resultRange[i, y + Config.StockCount * 3] = stock.MaxStock;
						y++;
					}
					firstIteration = false;
					i++;
				}
				// Выводим перемещение на лист
				Range[Cells[curentRow, 1], Cells[curentRow + supplierOrder.Count, ItemParametrsCount + Config.StockCount * 4]].Value2 = resultRange;
				// Применяем стили и форматирование
				Range[Cells[curentRow, 1], Cells[curentRow, ItemParametrsCount + Config.StockCount * 4]].Style = HeaderNameStyle;
				Range[Cells[curentRow, ItemParametrsCount], Cells[curentRow + supplierOrder.Count, ItemParametrsCount]].Style = CountNameStyle;
				Range[Cells[curentRow, 1], Cells[curentRow + supplierOrder.Count, ItemParametrsCount]].NumberFormat = NumberFormatText;
				// Границы колонок
				Range[Cells[curentRow + 1, ItemParametrsCount + 1], Cells[curentRow + supplierOrder.Count, ItemParametrsCount + Config.StockCount]].Borders(XlEdgeRight).Weight = XlThin;
				Range[Cells[curentRow + 1, ItemParametrsCount + 1], Cells[curentRow + supplierOrder.Count, ItemParametrsCount + Config.StockCount * 2]].Borders(XlEdgeRight).Weight = XlThin;
				Range[Cells[curentRow + 1, ItemParametrsCount + 1], Cells[curentRow + supplierOrder.Count, ItemParametrsCount + Config.StockCount * 3]].Borders(XlEdgeRight).Weight = XlThin;
				Range[Cells[curentRow + 1, ItemParametrsCount + 1], Cells[curentRow + supplierOrder.Count, ItemParametrsCount + Config.StockCount * 4]].Borders(XlEdgeRight).Weight = XlThin;

				curentRow += supplierOrder.Count + 1;
			}
		}
	}
}
