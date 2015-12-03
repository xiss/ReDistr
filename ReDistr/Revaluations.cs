﻿using System;
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
	public partial class Revaluations
	{
		private void Лист6_Startup(object sender, System.EventArgs e)
		{
		}

		private void Лист6_Shutdown(object sender, System.EventArgs e)
		{
		}

		#region Код, созданный конструктором VSTO

		/// <summary>
		/// Обязательный метод для поддержки конструктора - не изменяйте
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(Лист6_Startup);
			this.Shutdown += new System.EventHandler(Лист6_Shutdown);
		}

		#endregion


		// Определяем настройки
		private const int StartRow = 2;
		private const int ItemParametrsCount = 16;
		private const Excel.XlBordersIndex XlEdgeRight = Excel.XlBordersIndex.xlEdgeRight;
		private const Excel.XlBorderWeight XlThin = Excel.XlBorderWeight.xlThin;
		private const string CountNameStyle = "Хороший";
		private const string HeaderNameStyle = "Заголовок 1";
		private const string NumberFormatText = "@";

		// Выводит на лист заказы сгруппированные по поставщику
		public void FillList(List<Revaluation> revaluations)
		{
			// Очищаем лист
			// TODO /5 Подумать как сделать это проще, сейчас становятся активными лишние ячейки
			Range["A2:Z1500"].Clear();

			var curentRow = StartRow;

			var resultRange = new dynamic[revaluations.Count + 1, ItemParametrsCount];
			var i = 0;
			// Добавляем ЗЧ в массив 
			foreach (var revaluation in revaluations)
			{
				resultRange[i, 0] = revaluation.Item.Id1C;
				resultRange[i, 2] = revaluation.Item.Price;
				resultRange[i, 7] = revaluation.Item.Article;
				resultRange[i, 8] = revaluation.Item.Name;
				resultRange[i, 9] = revaluation.Item.Manufacturer;
				resultRange[i, 10] = revaluation.Item.StorageCategory;
				resultRange[i, 11] = revaluation.Item.CostPrice;
				resultRange[i, 12] = revaluation.NewPrice;
				resultRange[i, 13] = revaluation.Competitor.Count;
				resultRange[i, 14] = revaluation.Competitor.DeliveryTime;
				resultRange[i, 15] = revaluation.Competitor.Id;
				i++;
			}
			// Выводим перемещение на лист
			Range[Cells[curentRow, 1], Cells[curentRow + revaluations.Count, ItemParametrsCount]].Value2 = resultRange;

			// Применяем стили и форматирование
			/*Range[Cells[curentRow, 1], Cells[curentRow, ItemParametrsCount + Config.StockCount * 4]].Style = HeaderNameStyle;
			Range[Cells[curentRow, ItemParametrsCount], Cells[curentRow + supplierOrder.Count, ItemParametrsCount]].Style = CountNameStyle;
			Range[Cells[curentRow, 1], Cells[curentRow + supplierOrder.Count, ItemParametrsCount]].NumberFormat = NumberFormatText;
			// Границы колонок
			Range[Cells[curentRow + 1, ItemParametrsCount + 1], Cells[curentRow + supplierOrder.Count, ItemParametrsCount + Config.StockCount]].Borders(XlEdgeRight).Weight = XlThin;
			Range[Cells[curentRow + 1, ItemParametrsCount + 1], Cells[curentRow + supplierOrder.Count, ItemParametrsCount + Config.StockCount * 2]].Borders(XlEdgeRight).Weight = XlThin;
			Range[Cells[curentRow + 1, ItemParametrsCount + 1], Cells[curentRow + supplierOrder.Count, ItemParametrsCount + Config.StockCount * 3]].Borders(XlEdgeRight).Weight = XlThin;
			Range[Cells[curentRow + 1, ItemParametrsCount + 1], Cells[curentRow + supplierOrder.Count, ItemParametrsCount + Config.StockCount * 4]].Borders(XlEdgeRight).Weight = XlThin;
*/


		}
	}
}