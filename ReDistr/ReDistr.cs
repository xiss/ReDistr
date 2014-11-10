﻿using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using System.Text;

namespace ReDistr
{
	class ReDistr
	{
		private const int RowFirstFillNumber = 2;
		public static Control Control { get; set; }

	
		// Метод заполняет лист Test данными из словоря запчастей
		public static void FillTestList(Dictionary<string, Item> items)
		{
			Control.Application.Worksheets[3].Cells.Clear();
			// Заполняем заголовки
			Control.Application.Worksheets[3].Range["A1"].Value = "Id1C";
			Control.Application.Worksheets[3].Range["B1"].Value = "Name";
			Control.Application.Worksheets[3].Range["C1"].Value = "Article";
			Control.Application.Worksheets[3].Range["D1"].Value = "Manufacturer";
			Control.Application.Worksheets[3].Range["E1"].Value = "StorageCategory";
			Control.Application.Worksheets[3].Range["F1"].Value = "inBundle";
			Control.Application.Worksheets[3].Range["G1"].Value = "inKit";
			
			var curentRow = RowFirstFillNumber;
			foreach (KeyValuePair<string, Item> item in items)
			{
				// TODO нужно както по другому к листу обращаться
				Control.Application.Worksheets[3].Range["A" + curentRow].Value = item.Value.Id1C;
				Control.Application.Worksheets[3].Range["B" + curentRow].Value = item.Value.Name;
				Control.Application.Worksheets[3].Range["C" + curentRow].Value = item.Value.Article;
				Control.Application.Worksheets[3].Range["D" + curentRow].Value = item.Value.Manufacturer;
				Control.Application.Worksheets[3].Range["E" + curentRow].Value = item.Value.StorageCategory;
				Control.Application.Worksheets[3].Range["F" + curentRow].Value = item.Value.inBundle;
				Control.Application.Worksheets[3].Range["G" + curentRow].Value = item.Value.inKit;
				curentRow++;
			}

		}
	}
}