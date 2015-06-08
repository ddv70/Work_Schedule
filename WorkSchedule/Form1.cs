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
        public class Boundary
        {
            public class Pos
            {
                public int Begin { get; set; }
                public int End { get; set; }
            }
            Pos Date { get; set; }
            Pos DrName { get; set; }
            List<Pos> Department;
        }

        public Form1()
        {
            InitializeComponent();
            AllocConsole();
            LoadExcelFile();
        }

        HSSFWorkbook m_hssfwb;
        ISheet m_sheet;
        private void LoadExcelFile()
        {

            using (FileStream file = new FileStream("班表104-05.xls", FileMode.Open, FileAccess.Read))
            {
                m_hssfwb = new HSSFWorkbook(file);
            }

            m_sheet = m_hssfwb.GetSheetAt(0);
            for (int row = 0; row <= m_sheet.GetRow(0).LastCellNum; row++)
            {
                if (m_sheet.GetRow(1).GetCell(row) != null) //null is when the row only contains empty cells 
                {
                    Console.WriteLine(row.ToString()+" " + m_sheet.GetRow(1).GetCell(row));
                }
            }
        }
        private void DetectBoundary()
        {

        }
    }
}
