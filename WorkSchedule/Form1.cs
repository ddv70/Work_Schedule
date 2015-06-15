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

        public class DoctorWork
        {
            public string date { get; set; }
            public string DrName { get; set; }
            public string WorkType { get; set; }
            public string Department { get; set; }
        }

        public Form1()
        {
            InitializeComponent();
            AllocConsole();
            LoadExcelFile();
            DetectBoundary();
            FindDrOfType();
            ParseWorkingType();
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

        private string GetDrNameByPos(int pos)
        {
            string name = "";
            int begin = 9;
            for (int i = begin; i < begin + 3; i++)
                name += m_sheet.GetRow(i).GetCell(pos).StringCellValue;

            return name;
        }

        private List<DoctorWork> m_drworkintfo = new List<DoctorWork>();
        private void FindDrOfType()
        {
            for (int index_day = 14; index_day < 45; index_day++)
            {
                string date = m_sheet.GetRow(index_day).GetCell(0).NumericCellValue.ToString();
                Console.WriteLine("Date:{0}", date);
                for(int index_boundary=0;index_boundary<m_boundary.Department.Count-1;index_boundary++)
                {
                    Boundary.Pos p = m_boundary.Department[index_boundary];
                    for (int pos = p.Begin; pos <= p.End; pos++)
                    {
                        if (m_sheet.GetRow(index_day).GetCell(pos) != null)
                        {
                            
                            if (m_sheet.GetRow(index_day).GetCell(pos).StringCellValue != "" && m_sheet.GetRow(index_day).GetCell(pos).StringCellValue != "休")
                            {
                                DoctorWork dw = new DoctorWork();
                                string drName = "";
                                string type = m_sheet.GetRow(index_day).GetCell(pos).StringCellValue;
                                string depart = m_sheet.GetRow(1).GetCell(p.Begin).StringCellValue;
                                string name = GetDrNameByPos(pos);
                                if(!type.Contains("休"))
                                {
                                    dw.date = date;
                                    dw.DrName = name;
                                    dw.WorkType = type;
                                    dw.Department = depart;
                                    m_drworkintfo.Add(dw);
                                    //Console.WriteLine("Dr.:" + name + " type:" + type + " Depart:" + depart);
                                }
                            }
                        }
                    }
                }
                //Console.WriteLine();
            }
        }

        private void ParseWorkingType()
        {
            
            foreach(DoctorWork dw in m_drworkintfo)
            {
                Console.WriteLine("Date:{0},Dr.:{1}, Type:{2}, Depart:{3}",
                    dw.date,
                    dw.DrName,
                    dw.WorkType,
                    dw.Department
                    );
            }
        }
    }
}
