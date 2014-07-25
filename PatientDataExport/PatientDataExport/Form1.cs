using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

//include
using PatientDataExport.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace PatientDataExport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //患者的编号
            int peoplecount = 0;
            //设置要查询的时间
            DateTime startDate ;
            startDate = Convert.ToDateTime("2013-01-01 00:00:00");
            DateTime endDate;
            endDate = Convert.ToDateTime("2013-12-31 00:00:00");
            //文件的存储路径
            String FilePath;
            OpenFileDialog myFileDialog = new OpenFileDialog();
            myFileDialog.Filter = "Excel|*.xls";
            if (myFileDialog.ShowDialog() == DialogResult.OK)
            {
                FilePath = myFileDialog.FileName;
                MessageBox.Show(FilePath);
                Excel.Application myExcel = new Excel.Application();
                myExcel.Visible = true;
                Excel.Workbook myWorkbook = myExcel.Workbooks.Open(FilePath);
                Excel.Worksheet myWorkSheet = myWorkbook.Worksheets[1];
                medbaseEntities myMedBaseEntities = new medbaseEntities();
                //查询所有的待查询时间段内检查的患者
                var ExportResult = from s1 in myMedBaseEntities.hcheckmemb where s1.checkdate > startDate && s1.checkdate < endDate select s1;
                //遍历所有的患者
                foreach (var checkpatient in ExportResult)
                {
                    //遍历每一位患者
                    peoplecount++;
                    //患者的编号
                    myWorkSheet.Cells[peoplecount, 1] = peoplecount;
                    //姓名
                    myWorkSheet.Cells[peoplecount, 2] = checkpatient.a0101;
                    //性别
                    myWorkSheet.Cells[peoplecount, 3] = checkpatient.a0107;
                    //出生年月
                    myWorkSheet.Cells[peoplecount, 4] = 
                }
                myWorkbook.Save();
                myWorkbook.Close();

            }
        }

    }
}
