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
	public partial class ThisWorkbook
	{
		// Список ЗЧ
		public Dictionary<string, Item> items;

		

		private void ThisWorkbook_Startup(object sender, System.EventArgs e)
		{
			//TODO /0 Сделать здесь функцию проверяющую и если нужно создающую нужные каталоги

		}

		private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
		{
			
		}
		

		#region Код, созданный конструктором VSTO

		/// <summary>
		/// Обязательный метод для поддержки конструктора - не изменяйте
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(ThisWorkbook_Startup);
			this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
		}

		#endregion

	}

   

	
}
