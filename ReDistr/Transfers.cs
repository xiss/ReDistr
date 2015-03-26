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
	public partial class Transfers
	{
		private void Лист3_Startup(object sender, System.EventArgs e)
		{
		}

		private void Лист3_Shutdown(object sender, System.EventArgs e)
		{
		}

		#region Код, созданный конструктором VSTO

		/// <summary>
		/// Обязательный метод для поддержки конструктора - не изменяйте
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(Лист3_Startup);
			this.Shutdown += new System.EventHandler(Лист3_Shutdown);
		}

		#endregion

		// Определяем настройки
		private const int StartRow = 2;
		private const int ItemParametrsCount = 8;
		private const string TransferNameStyle = "Заголовок 1";
		// Выводит на лист перемещения из списка перемещений сгруппированные по направлениям
		public void FillList(List<Transfer> transfers)
		{
			// Список возможных направлений перемещений
			var unitedTransfers = ReDistr.GetPossibleTransfers(SimpleStockFactory.CurrentFactory.GetAllStocks()).ToList();
			var curentRow = StartRow;

			foreach (var unitedTransfer in unitedTransfers)
			{
				// Выбираем перемещения сгруппированные по направлению и объедененные по ЗЧ
				var transfer = unitedTransfer;
				var transferList = transfers.Where(
					transfer1 => transfer1.StockFrom.Name == transfer.StockFrom.Name && transfer1.StockTo.Name == transfer.StockTo.Name)
					.GroupBy(t => t.Item)
					.Select(tr => new Transfer()
					{
						StockFrom = transfer.StockFrom,
						StockTo = transfer.StockTo,
						Item = tr.First().Item,
						Count = tr.Sum(trs => trs.Count)

					}).ToList();

				var resultRange = new dynamic[transferList.Count + 1, ItemParametrsCount + Config.StockCount * 4];
				// Заполняем массив перемещениями
				resultRange[0, 0] = transferList.First().StockFrom.Name + " " + transferList.First().StockTo.Name;
				var i = 1;
				foreach (var curentTransfer in transferList)
				{
					resultRange[i, 0] = curentTransfer.Item.Id1C;
					resultRange[i, 1] = curentTransfer.Item.Name;
					resultRange[i, 2] = curentTransfer.Item.Article;
					resultRange[i, 3] = curentTransfer.Item.Manufacturer;
					resultRange[i, 4] = curentTransfer.Item.StorageCategory;
					resultRange[i, 5] = curentTransfer.Item.InBundle;
					resultRange[i, 6] = curentTransfer.Item.InKit;
					resultRange[i, 7] = curentTransfer.Count;
					// Добавляем информацию по складам
					var y = ItemParametrsCount;
					foreach (var stock in curentTransfer.Item.Stocks.OrderBy(t => t.Priority))
					{
						resultRange[i, y] = stock.Count;
						resultRange[i, y + Config.StockCount] = stock.SelingsCount;
						resultRange[i, y + Config.StockCount * 2] = stock.MinStock;
						resultRange[i, y + Config.StockCount * 3] = stock.MaxStock;
						y++;
					}
					i++;
				}
				// Выводим перемещение на лист
				Range[Cells[curentRow, 1], Cells[curentRow + transferList.Count, ItemParametrsCount + Config.StockCount * 4]].Value2 = resultRange;
				// Применяем стиль к заголовку
				Range[Cells[curentRow, 1], Cells[curentRow, ItemParametrsCount + Config.StockCount * 4]].Style = TransferNameStyle;

				curentRow += transferList.Count + 1;
			}
		}
	}
}
