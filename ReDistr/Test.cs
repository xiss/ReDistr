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

		private const int RowFirstFillNumber = 2;
		private const int ColumnFirstFillNumber = 8;
		public static Control Control { get; set; }



		// Метод заполняет лист Test данными из словоря запчастей
		// TODO Тупой метод, переделать
		public static void FillTestList(Dictionary<string, Item> items)
		{
			var curentRow = RowFirstFillNumber;
			var curentColumn = ColumnFirstFillNumber;

			Control.Application.Worksheets[3].Cells.ClearContents();

			// Заполняем заголовки
			Control.Application.Worksheets[3].Range["A1"].Value = "Id1C";
			Control.Application.Worksheets[3].Range["B1"].Value = "Name";
			Control.Application.Worksheets[3].Range["C1"].Value = "Article";
			Control.Application.Worksheets[3].Range["D1"].Value = "Manufacturer";
			Control.Application.Worksheets[3].Range["E1"].Value = "StorageCategory";
			Control.Application.Worksheets[3].Range["F1"].Value = "inBundle";
			Control.Application.Worksheets[3].Range["G1"].Value = "inKit";

			foreach (var item in items)
			{
				foreach (var stock in item.Value.Stocks)
				{
					Control.Application.Worksheets[3].Cells[1, curentColumn] = stock.Name;
					curentColumn = curentColumn + 3;

				}
				curentColumn = ColumnFirstFillNumber;
				break;
			}

			foreach (KeyValuePair<string, Item> item in items)
			{
				// Выводим информацию по ЗЧ
				// TODO нужно както по другому к листу обращаться
				Control.Application.Worksheets[3].Range["A" + curentRow].Value = item.Value.Id1C;
				Control.Application.Worksheets[3].Range["B" + curentRow].Value = item.Value.Name;
				Control.Application.Worksheets[3].Range["C" + curentRow].Value = item.Value.Article;
				Control.Application.Worksheets[3].Range["D" + curentRow].Value = item.Value.Manufacturer;
				Control.Application.Worksheets[3].Range["E" + curentRow].Value = item.Value.StorageCategory;
				Control.Application.Worksheets[3].Range["F" + curentRow].Value = item.Value.inBundle;
				Control.Application.Worksheets[3].Range["G" + curentRow].Value = item.Value.inKit;

				// Выводим информацию по складам
				foreach (var stock in item.Value.Stocks)
				{
					Control.Application.Worksheets[3].Cells[curentRow, curentColumn] = stock.InReserve;
					Control.Application.Worksheets[3].Cells[curentRow, curentColumn + 1] = stock.Priority;
					Control.Application.Worksheets[3].Cells[curentRow, curentColumn + 2] = stock.Count;
					curentColumn = curentColumn + 3;
				}

				curentColumn = ColumnFirstFillNumber;
				curentRow++;
			}
		}
	}
}
