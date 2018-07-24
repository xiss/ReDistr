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
		private const int ItemParametrsCount = 8;
		private const uint StockParametrsCount = 4;
		private const Excel.XlBordersIndex XlEdgeRight = Excel.XlBordersIndex.xlEdgeRight;
		private const Excel.XlBorderWeight XlThin = Excel.XlBorderWeight.xlThin;
		private const string CountNameStyle = "Хороший";
		private const string HeaderNameStyle = "Заголовок 1";
		private const string NumberFormatText = "@";
		// Строка с которой начинается заполненеие данными
		private const int ArrayRowFirstFillNumber = 3;
		// Первая колонка с которой выводятся параметры складов
		private const uint ArrayColumnFirstFillNumber = 8;


		// Выводит на лист заказы сгруппированные по поставщику
		public void FillList(List<Order> orders)
		{
			// Очищаем лист
			// TODO /5 Подумать как сделать это проще, сейчас становятся активными лишние ячейки
			//Cells.Clear();
			Range["A3:AB1500"].Clear();
			//Range["A3:E1500"].NumberFormat = NumberFormatText;

			var curentRow = ArrayRowFirstFillNumber;
			var resultRangeHeader = new dynamic[2, ItemParametrsCount + Config.Config.StockCount * 4];
			var curentColumn = ArrayColumnFirstFillNumber;
			var stockCount = Config.Config.StockCount;

			// Заполняем заголовки
			resultRangeHeader[1, 0] = "Id1C";
			resultRangeHeader[1, 1] = "Название";
			resultRangeHeader[1, 2] = "Артикул";
			resultRangeHeader[1, 3] = "Производитель";
			resultRangeHeader[1, 4] = "Кат. хранения";
			resultRangeHeader[1, 5] = "Упак.";
			resultRangeHeader[1, 6] = "Комплект";
			resultRangeHeader[1, 7] = "Кол.";

			// Выводим заголовки для параметров
			resultRangeHeader[0, curentColumn] = "Остатки";
			resultRangeHeader[0, curentColumn += stockCount] = "Продажи";
			resultRangeHeader[0, curentColumn += stockCount] = "Мин.";
			resultRangeHeader[0, curentColumn += stockCount] = "Макс.";

			// Выводим заголовки складов
			curentColumn = ArrayColumnFirstFillNumber;
			for (var i = 0; i < StockParametrsCount; i++)
			{
				foreach (var stock in orders.First().Item.Stocks)
				{
					//TODO /2 добавить короткое имя для склада
					resultRangeHeader[1, curentColumn] = stock.Name.Substring(0, 1);
					curentColumn++;
				}
			}
			// Заголовки выводим на лист
			Range[Cells[1, 1], Cells[2, ItemParametrsCount + Config.Config.StockCount * 4]].Value2 = resultRangeHeader;

			// Перебираем список поставщиков
			foreach (var supplier in Config.Config.ListSuppliers)
			{
				// Выбираем ЗЧ с данным поставщиком
				var supplierOrder = orders.Where(order => order.Item.Supplier == supplier).ToList();
				var resultRange = new dynamic[supplierOrder.Count + 1, ItemParametrsCount + Config.Config.StockCount * 4];
				resultRange[0, 0] = supplier;
				var i = 1;
				// Добавляем ЗЧ в массив 
				foreach (var order in supplierOrder)
				{
					resultRange[i, 0] = order.Item.Id1C;
					resultRange[i, 1] = order.Item.Name;
					resultRange[i, 2] = order.Item.Article;
					resultRange[i, 3] = order.Item.Manufacturer;
					resultRange[i, 4] = order.Item.StorageCategory;
					resultRange[i, 5] = order.Item.InBundle;
					resultRange[i, 6] = order.Item.InKit;
					resultRange[i, 7] = order.Count;

					// Добавляем информацию по складам
					var y = ItemParametrsCount;
					foreach (var stock in order.Item.Stocks)
					{
						resultRange[i, y] = stock.CountOrigin;
						resultRange[i, y + Config.Config.StockCount] = stock.CountSelings;
						resultRange[i, y + Config.Config.StockCount * 2] = stock.MinStock;
						resultRange[i, y + Config.Config.StockCount * 3] = stock.MaxStock;
						y++;
					}
					i++;
				}
				
				Range[Cells[curentRow, 1], Cells[curentRow + supplierOrder.Count, ItemParametrsCount]].NumberFormat = NumberFormatText;
				// Данные
				Range[Cells[curentRow, 1], Cells[curentRow + supplierOrder.Count, ItemParametrsCount + Config.Config.StockCount * 4]].Value2 = resultRange;
				// Применяем стили и форматирование
				Range[Cells[curentRow, 1], Cells[curentRow, ItemParametrsCount + Config.Config.StockCount * 4]].Style = HeaderNameStyle;
				Range[Cells[curentRow, ItemParametrsCount], Cells[curentRow + supplierOrder.Count, ItemParametrsCount]].Style = CountNameStyle;
				// Границы колонок
				Range[Cells[curentRow + 1, ItemParametrsCount + 1], Cells[curentRow + supplierOrder.Count, ItemParametrsCount + Config.Config.StockCount]].Borders(XlEdgeRight).Weight = XlThin;
				Range[Cells[curentRow + 1, ItemParametrsCount + 1], Cells[curentRow + supplierOrder.Count, ItemParametrsCount + Config.Config.StockCount * 2]].Borders(XlEdgeRight).Weight = XlThin;
				Range[Cells[curentRow + 1, ItemParametrsCount + 1], Cells[curentRow + supplierOrder.Count, ItemParametrsCount + Config.Config.StockCount * 3]].Borders(XlEdgeRight).Weight = XlThin;
				Range[Cells[curentRow + 1, ItemParametrsCount + 1], Cells[curentRow + supplierOrder.Count, ItemParametrsCount + Config.Config.StockCount * 4]].Borders(XlEdgeRight).Weight = XlThin;

				curentRow += supplierOrder.Count + 1;
			}
		}
	}
}
