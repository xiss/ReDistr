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
using System.Xml.Serialization;
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

		// Строка с которой начинается заполненеие данными
		private const uint ArrayRowFirstFillNumber = 2;
		// Первая колонка с которой выводятся параметры складов
		private const uint ArrayColumnFirstFillNumber = 7;
		// Количество параметров для склада
		private const uint StockParametrsCount = 8;

		// Метод заполняет лист Test данными из словоря запчастей
		public void FillListStocks(Dictionary<string, Item> items)
		{
			var curentRow = ArrayRowFirstFillNumber;
			var curentColumn = ArrayColumnFirstFillNumber;
			var stockCount = Config.StockCount;
			// 7 колонок под описание товара 8 колонок под параметры склада
			var resultRange = new dynamic[items.Count + 2, ArrayColumnFirstFillNumber + Config.StockCount * StockParametrsCount];
			Cells.ClearContents();

			// Заполняем заголовки ЗЧ
			resultRange[0, 0] = "Id1C";
			resultRange[0, 1] = "Name";
			resultRange[0, 2] = "Article";
			resultRange[0, 3] = "Manufacturer";
			resultRange[0, 4] = "StorageCategory";
			resultRange[0, 5] = "inBundle";
			resultRange[0, 6] = "inKit";

			// Выводим заголовки для параметров
			resultRange[0, curentColumn] = "Count";
			resultRange[0, curentColumn += stockCount] = "InReserve";
			resultRange[0, curentColumn += stockCount] = "SelingsCount";
			resultRange[0, curentColumn += stockCount] = "SailPersent";
			resultRange[0, curentColumn += stockCount] = "MinStock";
			resultRange[0, curentColumn += stockCount] = "MaxStock";
			//resultRange[0, curentColumn += stockCount] = "FreeStock";
			//resultRange[0, curentColumn += stockCount] = "Need";
			resultRange[0, curentColumn += stockCount] = "Priority";
			resultRange[0, curentColumn += stockCount] = "Exclude";

			// Выводим заголовки складов
			curentColumn = ArrayColumnFirstFillNumber;
			for (var i = 0; i < StockParametrsCount; i++)
			{
				foreach (var stock in items.First().Value.Stocks)
				{
					//TODO добавить короткое имя для склада
					resultRange[1, curentColumn] = stock.Name.Substring(0, 1);
					curentColumn++;
				}
			}
			foreach (var item in items)
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
				uint curentStock = 1;
				foreach (var stock in item.Value.Stocks)
				{
					resultRange[curentRow, curentColumn] = stock.Count;
					resultRange[curentRow, curentColumn += stockCount] = stock.InReserve;
					resultRange[curentRow, curentColumn += stockCount] = stock.SelingsCount;
					resultRange[curentRow, curentColumn += stockCount] = stock.SailPersent;
					resultRange[curentRow, curentColumn += stockCount] = stock.MinStock;
					resultRange[curentRow, curentColumn += stockCount] = stock.MaxStock;
					//resultRange[curentRow, curentColumn += stockCount] = stock.FreeStock;
					//resultRange[curentRow, curentColumn += stockCount] = stock.Need;
					resultRange[curentRow, curentColumn += stockCount] = stock.Priority;
					resultRange[curentRow, curentColumn += stockCount] = stock.ExcludeFromMoovings;

					curentColumn = ArrayColumnFirstFillNumber;
					curentColumn += curentStock;
					curentStock++;
				}
				curentRow++;
			}

			// Выводим результат на лист
			Range[Cells[1, 1], Cells[items.Count + 2, 7 + StockParametrsCount * Config.StockCount]].Value2 = resultRange;
			//return resultRange;
		}

		// Выводит информацию о перемещениях на лист
		public void FillListTransfers(List<Transfer> transfers, Dictionary<string, Item> items)
		{
			// Первая колонка после складов
			var FirstColumn = ArrayColumnFirstFillNumber + Config.StockCount * StockParametrsCount + 1;

			// Список возможных направлений перемещений
			var possibleTransfers = ReDistr.GetPossibleTransfers(SimpleStockFactory.CurrentFactory.GetAllStocks()).ToList();

			// Итоговый массив для вывода на лист
			var resultRange = new dynamic[items.Count + 1, possibleTransfers.Count];

			// Строка в массиве
			var curentArrayRow = 0;

			// Выводим заголовки для пеермещений
			for (int j = 0; j < possibleTransfers.Count; j++)
			{
				resultRange[curentArrayRow, j] = possibleTransfers[j].StockFrom.Name.Substring(0, 1) +
												 possibleTransfers[j].StockTo.Name.Substring(0, 1);
			}
			curentArrayRow++;

			// Выводим перемещения
			var curentRow = ArrayRowFirstFillNumber + 1;
			while (Range["A" + curentRow].Value2 != null)
			{
				string CurentId1C = Range["A" + curentRow].Value2;
				for (int j = 0; j < possibleTransfers.Count; j++)
				{
					var query = from transfer in transfers
								where transfer.StockFrom.Name == possibleTransfers[j].StockFrom.Name &&
							  transfer.StockTo.Name == possibleTransfers[j].StockTo.Name &&
							  transfer.Item.Id1C == CurentId1C
								select transfer;

					resultRange[curentArrayRow, j] = query.Sum(transfer => transfer.Count);
				}
				curentArrayRow++;
				curentRow++;
			}

			// Выводим результат на лист
			Range[Cells[2, FirstColumn], Cells[items.Count + 1, FirstColumn + possibleTransfers.Count - 1]].Value2 = resultRange;
		}
	}
}
