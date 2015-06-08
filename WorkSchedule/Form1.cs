using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Runtime.InteropServices;

namespace WorkSchedule
{
    public partial class Form1 : Form
    {
        [DllImport("kernel32.dll",
        EntryPoint = "AllocConsole",
        SetLastError = true,
        CharSet = CharSet.Auto,
        CallingConvention = CallingConvention.StdCall)]
        private static extern int AllocConsole();

        public Form1()
        {
            InitializeComponent();
            AllocConsole();
            LoadExcelFile();
        }

        HSSFWorkbook m_hssfwb;
        ISheet sheet;
        private void LoadExcelFile()
        {

            using (FileStream file = new FileStream("班表104-05.xls", FileMode.Open, FileAccess.Read))
            {
                m_hssfwb = new HSSFWorkbook(file);
            }

            ISheet sheet = m_hssfwb.GetSheetAt(0);
            for (int row = 0; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                {
                    Console.WriteLine(sheet.GetRow(row).GetCell(0));
                }
            }
        }
    }
}
