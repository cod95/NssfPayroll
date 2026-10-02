using System;
using System.Windows.Forms;

namespace NssfPayroll
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            ApplicationConfiguration.Initialize();
            ExcelPrinter.CleanOldTempFiles();
            Application.Run(new MainForm());
        }
    }
}
