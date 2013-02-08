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
           
            //string oldval = " ";
            //string mapval = " ";
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
            for (int createsheet_i = 1; createsheet_i <= 100; createsheet_i++)    //สร้าง ชีท ใหม่ โดยใช้ข้อมู,จาก old sheet
            {
                Range old_valuerow = old.Cells[createsheet_i, 5];

                for (int createsheet = 1; createsheet <= 10; createsheet++)
                {
                    Range old_value = old.Cells[createsheet_i, createsheet];    //ได้ ค่าจาก ชีท เดิม มา
                    new_sheetbook.Cells[createsheet_i, createsheet] = old_value.Value;
                    richTextBox3.Text = "created at row [" + createsheet_i + "," + createsheet + "].....\n";
                }

                
            }
            richTextBox3.Text += "New sheet are created .....\n";

            string countlog = " ";
            for (int i = 1; i <= 400; i++)
            {
                Range mapping_cell = mapping.Cells[i, 1];   //เรียกค่า ใน mapping sheet แถวที่ i คอลัมที่ 1
                richTextBox1.Text = "Compare at row " + i+"\n";
                string mapping2 = " " + mapping_cell.Value;

                for (int x = 1; x <= 10; x++) {        //กำหนดจำนวนของ แถวที่ให้ค้น หาใน sheet old
                    richTextBox2.Text = "compare at row " + x + "\n";
                    richTextBox3.Text += countlog;

                    Range old_cell = old.Cells[i,5];
                    
                        string old2 = " " + old_cell.Value;
                    
                    
                  //  string mappingstr = mapping_cell.Value.ToString();
                   // string oldstr = old_cell.Value.ToString();
                    
                    if ("1A00000" == mapping2.Trim())
                    {
                        new_sheetbook.Cells[i, 5] = mapping.Cells[i, 6];
                        new_sheetbook.Cells[i, 6] = mapping.Cells[i, 8];
                      // countlog += "Edit at Row " + i + "Ma pping = " + mapping_cell.Value + ", Source file = " + old_cell.Value + " .\n";
                    }
                    old2 = old_cell.Value;
                   
                    
                
                }
                
            }
            richTextBox2.Text += "Compare Finished";
            richTextBox1.Text += "Compare Finished";

        }
    }
}
