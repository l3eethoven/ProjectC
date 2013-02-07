using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace CAT_Element
{
    public partial class Form1 : Form
    {
        Microsoft.Office.Interop.Excel.Application m_app;

        public Form1()
        {
            
            InitializeComponent();
            m_app = new Microsoft.Office.Interop.Excel.Application();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog browseM = new OpenFileDialog();
            browseM.Title = "Select your mapping file";
            browseM.InitialDirectory = @"c:\";
            browseM.Filter = "Excel 2007-2010 (*.xlsx)|*.xlsx|Excel 2003(*.xls)|*.xls|All files(*.*)|*.*";
            browseM.FilterIndex = 1;
            browseM.RestoreDirectory = true;
            if (browseM.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = browseM.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog browse = new OpenFileDialog();
            browse.Title = "Select your old file";
            browse.InitialDirectory = @"c:\";
            browse.Filter = "Excel 2007-2010 (*.xlsx)|*.xlsx|Excel 2003(*.xls)|*.xls|All files(*.*)|*.*";
            browse.FilterIndex = 1;
            browse.RestoreDirectory = true;
            if (browse.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = browse.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int cellnone = 1;
            int cellnonerow = 1;
            Workbook Mapping_workbook = m_app.Workbooks.Open(textBox1.Text,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
                , Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            if (Mapping_workbook == null)
                return;
           
            Workbook Old_workbook = m_app.Workbooks.Open(textBox2.Text,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
                , Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            if (Old_workbook == null)
                return;

            Worksheet mapping = Mapping_workbook.Sheets[1];
            Worksheet old = Old_workbook.Sheets[2];
            Worksheet new_sheetbook = Old_workbook.Sheets[Old_workbook.Sheets.Count];

            richTextBox3.Text += "Creating New sheet .....\n";
            for (int createsheet_i = 1; createsheet_i <= 40000; createsheet_i++)    //สร้าง ชีท ใหม่ โดยใช้ข้อมู,จาก old sheet
            {
                Range old_valuerow = old.Cells[createsheet_i, 5];

                for (int createsheet = 1; createsheet <= 100; createsheet++)
                {
                    Range old_value = old.Cells[createsheet_i, createsheet];    //ได้ ค่าจาก ชีท เดิม มา
                    new_sheetbook.Cells[createsheet_i, createsheet] = old_value.Value;

                    if (old_value.Value == null) {
                        cellnone++;
                        if (cellnone >= 20)
                        {
                            createsheet = 101;
                        }
                        else cellnone = 1;
                    }
                }

                if (old_valuerow.Value == null)
                {
                    cellnone++;
                    if (cellnonerow >= 50)
                    {
                        createsheet_i = 40001;
                    }
                    else cellnonerow = 1;
                }
            }
            richTextBox3.Text += "New sheet are created .....\n";
            


            cellnone = 1;

            for (int i = 1; i <= 40000; i++)
            {
                Range mapping_cell = mapping.Cells[i, 1];   //เรียกค่า ใน mapping sheet แถวที่ i คอลัมที่ 1
                richTextBox1.Text = "Compare at row " + i+"\n";

                for (int x = 1; x <= 40000; x++) {        //กำหนดจำนวนของ แถวที่ให้ค้น หาใน sheet old
                    richTextBox2.Text += "compare at row " + x + "\n";
                    Range old_cell = old.Cells[i,5];
                    
                    if (mapping_cell.Value == old_cell.Value) {
                        new_sheetbook.Cells[i, 5] = mapping.Cells[i, 6];
                        new_sheetbook.Cells[i, 6] = mapping.Cells[i, 8];
                    }
                    
                
                }





                if (mapping_cell.Value == null)               //ส่วนตรวจสอบว่า cell ว่าง หรือเปล่า
                {
                    if (cellnone >= 100)
                    {                   //หาก Cell ว่างเกิน 100 ช่อง ให้ ออกจากลูป โดยการเซ็ต i = 50000         
                        i = 50000;
                    }
                    
                    cellnone++;                             //หาก cell ว่างและ ยังติดต่อกันไม่ถึง 100 ช่อง ให้เพิ่ม ค่า cell none อีก 1

                }
                else {
                    cellnone = 1;                           // ถ้า หาก ว่า cell ไม่ว่าง ให้กำหนด ค่า cellnone ที่ 1 ใหม่

                }




            }

        }
    }
}
