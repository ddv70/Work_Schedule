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
            public Pos Date { get; set; }
            public Pos DrName { get; set; }
            public List<Pos> Department;
        }

        public Form1()
        {
            InitializeComponent();
            AllocConsole();
            LoadExcelFile();
            DetectBoundary();
            FindDrOfType();
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
        Boundary m_boundary = new Boundary();
        
        private void DetectBoundary()
        {
            m_boundary.Department = new List<Boundary.Pos>();

            List<int> tmp_list = new List<int>();
            int tmp_bg = 0;
            int tmp_end = 0; 
            for (int index = 2; index < m_sheet.GetRow(1).LastCellNum; index++)
            {
                if (m_sheet.GetRow(1).GetCell(index) != null)
                {
                    if (m_sheet.GetRow(1).GetCell(index).StringCellValue != "")
                        tmp_list.Add(index);
                }
                
            }
            for (int i = 0; i < tmp_list.Count;i++ )
            {
                Boundary.Pos p = new Boundary.Pos();
                p.Begin = tmp_list[i];
                if (i == tmp_list.Count - 1)
                    p.End = m_sheet.GetRow(1).LastCellNum;
                else
                    p.End = tmp_list[i + 1] - 1;

                m_boundary.Department.Add(p);

            }
            //test
            
            foreach(Boundary.Pos p in m_boundary.Department )
            {
                Console.WriteLine("pos b:"+p.Begin+" e:"+p.End);
            }
        }

        private void FindDrOfType()
        {
            for (int index_day = 14; index_day < 45; index_day++)
            {
                foreach (Boundary.Pos p in m_boundary.Department)
                {
                    for (int pos = p.Begin; pos <= p.End; pos++)
                    {
                        if (m_sheet.GetRow(index_day).GetCell(pos).StringCellValue != "")
                        {
                            string drName = "";
                            string type = m_sheet.GetRow(index_day).GetCell(pos).StringCellValue;
                            Console.WriteLine("pos:"+pos+" type:"+type);
                        }
                    }
                }
            }
        }
    }
}
