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
		private const uint ArrayColumnFirstFillNumber = 10;
		// Количество параметров для склада
		private const uint StockParametrsCount = 9;

		// Метод заполняет лист Test данными из словоря запчастей
		public void FillListStocks(Dictionary<string, Item> items)
		{
			var curentRow = ArrayRowFirstFillNumber;
			var curentColumn = ArrayColumnFirstFillNumber;
			var stockCount = Config.StockCount;
			// 7 колонок под описание товара 9 колонок под параметры склада
			var resultRange = new dynamic[items.Count + 2, ArrayColumnFirstFillNumber + Config.StockCount * StockParametrsCount];
			Cells.ClearContents();

			// Заполняем заголовки ЗЧ
			resultRange[0, 0] = "Id1C";
			resultRange[0, 1] = "Name";
			resultRange[0, 2] = "Article";
			resultRange[0, 3] = "Manufacturer";
			resultRange[0, 4] = "Supplier";
			resultRange[0, 5] = "StorageCat";
			resultRange[0, 6] = "inBundle";
			resultRange[0, 7] = "inKit";
			resultRange[0, 8] = "CostPrice";
			resultRange[0, 9] = "Price";

			// Выводим заголовки для параметров
			resultRange[0, curentColumn] = "Count";
			resultRange[0, curentColumn += stockCount] = "InReserve";
			resultRange[0, curentColumn += stockCount] = "SelCount";
			resultRange[0, curentColumn += stockCount] = "SailPers";
			resultRange[0, curentColumn += stockCount] = "MinStock";
			resultRange[0, curentColumn += stockCount] = "MaxStock";
			resultRange[0, curentColumn += stockCount] = "Priority";
			resultRange[0, curentColumn += stockCount] = "Exclude";
			resultRange[0, curentColumn += stockCount] = "ReqAv";

			// Выводим заголовки складов
			curentColumn = ArrayColumnFirstFillNumber;
			for (var i = 0; i < StockParametrsCount; i++)
			{
				foreach (var stock in items.First().Value.Stocks)
				{
					//TODO /2 добавить короткое имя для склада
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
				resultRange[curentRow, 4] = item.Value.Supplier;
				resultRange[curentRow, 5] = item.Value.StorageCategory;
				resultRange[curentRow, 6] = item.Value.InBundle;
				resultRange[curentRow, 7] = item.Value.InKit;
				resultRange[curentRow, 8] = item.Value.CostPrice;
				resultRange[curentRow, 9] = item.Value.Price;


				// Выводим информацию по складам
				curentColumn = ArrayColumnFirstFillNumber;
				uint curentStock = 1;
				foreach (var stock in item.Value.Stocks)
				{
					resultRange[curentRow, curentColumn] = stock.Count;
					resultRange[curentRow, curentColumn += stockCount] = stock.InReserve;
					resultRange[curentRow, curentColumn += stockCount] = stock.CountSelings;
					resultRange[curentRow, curentColumn += stockCount] = stock.SailPersent;
					resultRange[curentRow, curentColumn += stockCount] = stock.MinStock;
					resultRange[curentRow, curentColumn += stockCount] = stock.MaxStock;
					resultRange[curentRow, curentColumn += stockCount] = stock.Priority;
					resultRange[curentRow, curentColumn += stockCount] = ReDistr.BoolToInt(stock.ExcludeFromMoovings);
					resultRange[curentRow, curentColumn += stockCount] = ReDistr.BoolToInt(stock.RequiredAvailability);
					
					curentColumn = ArrayColumnFirstFillNumber;
					curentColumn += curentStock;
					curentStock++;
				}
				curentRow++;
			}

			// Выводим результат на лист
			Range[Cells[1, 1], Cells[items.Count + 2, ArrayColumnFirstFillNumber + StockParametrsCount * Config.StockCount]].Value2 = resultRange;
		}

		// Выводит информацию о перемещениях на лист
		public void FillListTransfers(List<Transfer> transfers, Dictionary<string, Item> items)
		{
			// Первая колонка после складов
			var firstColumn = ArrayColumnFirstFillNumber + Config.StockCount * StockParametrsCount + 1;

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
					// Получаем список перемещенн
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
			Range[Cells[ArrayRowFirstFillNumber, firstColumn], Cells[ArrayRowFirstFillNumber + items.Count, firstColumn + possibleTransfers.Count - 1]].Value2 = resultRange;
		}
	}
}
