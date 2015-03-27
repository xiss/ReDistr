﻿using System;
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
		private const int StartRow = 3;
		private const int ItemParametrsCount = 8;
		private const string TransferNameStyle = "Заголовок 1";
		// Выводит на лист перемещения из списка перемещений сгруппированные по направлениям
		public void FillList(List<Transfer> transfers)
		{
			// Список возможных направлений перемещений
			var unitedTransfers = Config.PossibleTransfers;
			var curentRow = StartRow;
			var firstIteration = true;
			var resultRangeHeader = new dynamic[1, ItemParametrsCount + Config.StockCount * 4 + Config.CountPossibleTransfers];

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

				var resultRange = new dynamic[transferList.Count + 1, ItemParametrsCount + Config.StockCount * 4 + Config.CountPossibleTransfers];
				// Заполняем массив перемещениями
				resultRange[0, 0] = transferList.First().StockFrom.Name + " => " + transferList.First().StockTo.Name;
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
						resultRange[i, y + Config.StockCount] = stock.SelingsCount;
						resultRange[i, y + Config.StockCount * 2] = stock.MinStock;
						resultRange[i, y + Config.StockCount * 3] = stock.MaxStock;
						y++;
					}
					firstIteration = false;
					// Добавляем информацию о перемещениях
					y += (int)Config.StockCount * 3;
					foreach (var possibleTransfer in unitedTransfers)
					{
						// Получаем список перемещений с данным направлением и запчастью
						var query = from transfer1 in transfers
									where transfer1.StockFrom.Name == possibleTransfer.StockFrom.Name &&
								  transfer1.StockTo.Name == possibleTransfer.StockTo.Name &&
								  transfer1.Item.Id1C == curentTransfer.Item.Id1C
									select transfer1;

						resultRange[i, y] = query.Sum(transfer1 => transfer1.Count);
						resultRangeHeader[0, y] = possibleTransfer.StockFrom.Name.Substring(0, 1) +
						                          possibleTransfer.StockTo.Name.Substring(0, 1);
						y++;
					}
					i++;
				}

				// Выводим перемещение на лист
				Range[Cells[curentRow, 1], Cells[curentRow + transferList.Count, ItemParametrsCount + Config.StockCount * 4 + Config.CountPossibleTransfers]].Value2 = resultRange;
				// Применяем стиль к заголовку
				Range[Cells[curentRow, 1], Cells[curentRow, ItemParametrsCount + Config.StockCount * 4 + Config.CountPossibleTransfers]].Style = TransferNameStyle;

				curentRow += transferList.Count + 1;
			}
			// Выводим заголовки
			Range[Cells[2, 1], Cells[2, ItemParametrsCount + Config.StockCount * 4 + Config.CountPossibleTransfers]].Value2 = resultRangeHeader;
		}
	}
}
