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
		private const int ItemParametrsCount = 20;
		private const Excel.XlBordersIndex XlEdgeRight = Excel.XlBordersIndex.xlEdgeRight;
		private const Excel.XlBorderWeight XlThin = Excel.XlBorderWeight.xlThin;
		private const string CountNameStyle = "Хороший";
		private const string HeaderNameStyle = "Заголовок 1";
		private const string NumberFormatText = "@";
		private const string NumberFormatPersent = "0%";

		// Выводит на лист заказы сгруппированные по поставщику
		public void FillList(List<Revaluation> revaluations)
		{
			// Очищаем лист
			// TODO /5 Подумать как сделать это проще, сейчас становятся активными лишние ячейки
			Range["A2:Z15000"].Clear();

			var resultRange = new dynamic[revaluations.Count + 1, ItemParametrsCount];
			var i = 0;
			// Добавляем ЗЧ в массив 
			foreach (var revaluation in revaluations)
			{
				resultRange[i, 0] = "'" + revaluation.Item.Id1C;
				resultRange[i, 1] = revaluation.Item.Name;
				resultRange[i, 2] = "'" + revaluation.Item.Article;
				resultRange[i, 3] = revaluation.Item.Manufacturer;
				resultRange[i, 4] = revaluation.Item.StorageCategory;
				resultRange[i, 5] = revaluation.Item.GetSumStocks();
				resultRange[i, 6] = revaluation.Item.Property;
				resultRange[i, 7] = revaluation.Item.Price;
				resultRange[i, 8] = revaluation.Item.GetAVGCostPrice();
				resultRange[i, 9] = revaluation.NewPrice;
				resultRange[i, 10] = revaluation.NewPrice - revaluation.Item.Price;
				resultRange[i, 11] = (revaluation.NewPrice - revaluation.Item.GetAVGCostPrice()) / revaluation.Item.GetAVGCostPrice();
				resultRange[i, 12] = revaluation.Item.GetPricePortalWithAdd();
				if (revaluation.Competitor != null)
				{
					resultRange[i, 13] = Math.Round(revaluation.Competitor.PriceWithAdd, 2);
					resultRange[i, 14] = revaluation.Competitor.Count;
					resultRange[i, 15] = revaluation.Competitor.DeliveryTime;
					resultRange[i, 16] = revaluation.Competitor.Id;
					resultRange[i, 17] = revaluation.Competitor.PositionNumber;
					resultRange[i, 18] = revaluation.Competitor.Region;
				}
				resultRange[i, 19] = revaluation.Note;
				i++;
			}
			// Выводим переоценку на лист
			Range[Cells[StartRow, 1], Cells[revaluations.Count, ItemParametrsCount]].Value2 = resultRange;

			// Применяем стили и форматирование
			Range["L2:L" + i].NumberFormat = NumberFormatPersent;
		}

		// Создает книгу с переоценкой
		public void MakeImportRevaluation()
		{
			var firstRow = StartRow;
			var bookName = DateTime.Now.ToShortDateString() + ".xls";
			// Определяем последнюю строку
			var i = 1;
			do
			{
				i++;
			}
			while (Range["A" + i].Value2 != null);
			var lastRow = i;

			var selectionRange = Application.Union(Range["A" + firstRow + ":" + "A" + lastRow], Range["F" + firstRow + ":" + "F" + lastRow], Range["J" + firstRow + ":" + "J" + lastRow]);
			ReDistr.MakeImpot1CBook(selectionRange, bookName, Config.FolderRevaluations);
		}
	}
}
