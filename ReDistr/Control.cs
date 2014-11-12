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
    public partial class Control
    {
        private void Лист1_Startup(object sender, System.EventArgs e)
        {
        }

        private void Лист1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Код, созданный конструктором VSTO

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.buttonGetMoving.Click += new System.EventHandler(this.buttonGetMoving_Click);
            this.Startup += new System.EventHandler(this.Лист1_Startup);
            this.Shutdown += new System.EventHandler(this.Лист1_Shutdown);

        }

        #endregion

        private void buttonGetMoving_Click(object sender, EventArgs e)
        {
            var parser = new Parser(this);
            var items = parser.Parse();
			Test.Control = this;
			Test.FillTestList(items);
        }
    }
}
