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
            //禁用开始处理按钮
            btn_beginProgress.Enabled = false;
            //患者的编号
            int peoplecount = 0;
            //设置要查询的时间
            DateTime startDate ;
            startDate = Convert.ToDateTime("2014-01-01 00:00:00");
            DateTime endDate;
            endDate = Convert.ToDateTime("2014-12-31 00:00:00");
            //文件的存储路径
            String FilePath = txtbox_FilePath.Text;
            Excel.Application myExcel = new Excel.Application();
            myExcel.Visible = false;
            Excel.Workbook myWorkbook = myExcel.Workbooks.Add(true);
            Excel.Worksheet myWorkSheet = myWorkbook.Worksheets[1];
            medbaseEntities myMedBaseEntities = new medbaseEntities();
            //查询所有的待查询时间段内检查的患者
            var ExportResult = from s1 in myMedBaseEntities.hcheckmemb where s1.checkdate > startDate && s1.checkdate < endDate select s1;
            //总数
            totalNum.Text = ExportResult.Count().ToString();
            //遍历所有的患者
            foreach (var checkpatient in ExportResult)
            {
                //遍历每一位患者
                peoplecount++;
                progressNum.Text = peoplecount.ToString();
                //患者的编号
                myWorkSheet.Cells[peoplecount, 1] = peoplecount;
                //姓名
                myWorkSheet.Cells[peoplecount, 2] = checkpatient.a0101;
                //性别
                myWorkSheet.Cells[peoplecount, 3] = checkpatient.a0107;
                //出生年月
                try
                {
                    var searchBirthday = (from s2 in myMedBaseEntities.hbasememb where s2.membcode == checkpatient.membcode  select s2).First();
                    //出生年月
                    myWorkSheet.Cells[peoplecount, 4] = searchBirthday.a0111.GetValueOrDefault().ToShortDateString();
                    //移动电话
                    myWorkSheet.Cells[peoplecount, 6] = searchBirthday.mobileno.ToString();
                    //保健类型
                    myWorkSheet.Cells[peoplecount, 7] = searchBirthday.a0704.ToString();
                    //保健证号
                    myWorkSheet.Cells[peoplecount, 8] = searchBirthday.medicareno.ToString();
                }
                catch
                {
                    //出生年月
                    myWorkSheet.Cells[peoplecount, 4] = "空白";
                    //移动电话
                    myWorkSheet.Cells[peoplecount, 6] = "空白";
                    //保健证号
                    myWorkSheet.Cells[peoplecount, 8] = "空白";
                }
                finally
                {
 
                }
                //工作单位
                myWorkSheet.Cells[peoplecount, 5] = checkpatient.b0105.ToString();
                //体检医院
                myWorkSheet.Cells[peoplecount, 9] = "天津医科大学总医院";
                //各个检查结果
                try
                {
                    var testResult = from s3 in myMedBaseEntities.hdatadeptest where checkpatient.checkcode == s3.checkcode  select s3;
                    foreach (var eachtest in testResult)
                    {
                        //身高
                        if (eachtest.testcode == "D4.0010")
                        {
                            myWorkSheet.Cells[peoplecount, 22] = eachtest.testresult; 
                        }
                        //体重
                        if (eachtest.testcode == "D4.0020")
                        {
                            myWorkSheet.Cells[peoplecount, 23] = eachtest.testresult;
                        }
                        //血压
                        if (eachtest.testcode == "D4.0040") 
                        {
                            try
                            {
                                string[] splitestring = eachtest.testresult.Split('/');
                                myWorkSheet.Cells[peoplecount, 24] = splitestring[0].ToString();
                                myWorkSheet.Cells[peoplecount, 25] = splitestring[1].ToString();
                            }
                            catch
                            {
                                myWorkSheet.Cells[peoplecount, 24] = "空白";
                                myWorkSheet.Cells[peoplecount, 25] = "空白";
                            }
                            finally
                            { 
                            }
                        }
                        //谷丙转氨酶（ALT）
                        if (eachtest.testcode == "Y1.0010")
                        {
                            myWorkSheet.Cells[peoplecount, 45] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 46] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 47] = eachtest.testhigher;
                        }
                        //总胆红素(TBIL)
                        if (eachtest.testcode == "Y1.0050")
                        {
                            myWorkSheet.Cells[peoplecount, 48] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 49] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 50] = eachtest.testhigher;
                        }
                        //直接胆红素(DBIL)
                        if (eachtest.testcode == "Y1.0060")
                        {
                            myWorkSheet.Cells[peoplecount, 51] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 52] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 53] = eachtest.testhigher;
                        }
                        //总蛋白(TP)
                        if (eachtest.testcode == "Y1.0080")
                        {
                            myWorkSheet.Cells[peoplecount, 54] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 55] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 56] = eachtest.testhigher;
                        }
                        //白蛋白(ALB)
                        if (eachtest.testcode == "Y1.0090")
                        {
                            myWorkSheet.Cells[peoplecount, 57] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 58] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 59] = eachtest.testhigher;
                        }
                        //球蛋白(GLB)
                        if (eachtest.testcode == "Y1.0100")
                        {
                            myWorkSheet.Cells[peoplecount, 60] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 61] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 62] = eachtest.testhigher;
                        }
                        //尿素氮(BUN)
                        if (eachtest.testcode == "Y2.0010")
                        {
                            myWorkSheet.Cells[peoplecount, 63] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 64] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 65] = eachtest.testhigher;
                        }
                        //肌酐(Cr)
                        if (eachtest.testcode == "Y2.0020")
                        {
                            myWorkSheet.Cells[peoplecount, 66] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 67] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 68] = eachtest.testhigher;
                        }
                        //尿酸(UA)
                        if (eachtest.testcode == "Y2.0030")
                        {
                            myWorkSheet.Cells[peoplecount, 69] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 70] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 71] = eachtest.testhigher;
                        }
                        //葡萄糖(GLU)
                        if (eachtest.testcode == "Y3.0010")
                        {
                            myWorkSheet.Cells[peoplecount, 72] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 73] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 74] = eachtest.testhigher;
                        }
                        //糖化血红蛋白
                        if (eachtest.testcode == "Y3.0030")
                        {
                            myWorkSheet.Cells[peoplecount, 75] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 76] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 77] = eachtest.testhigher;
                        }
                        //总胆固醇(CHOL)
                        if (eachtest.testcode == "Y4.0010")
                        {
                            myWorkSheet.Cells[peoplecount, 78] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 79] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 80] = eachtest.testhigher;
                        }
                        //甘油三酯(TG)
                        if (eachtest.testcode == "Y4.0020")
                        {
                            myWorkSheet.Cells[peoplecount, 78] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 79] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 80] = eachtest.testhigher;
                        }
                        //高密度脂蛋白
                        if (eachtest.testcode == "Y4.0030")
                        {
                            myWorkSheet.Cells[peoplecount, 84] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 85] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 86] = eachtest.testhigher;
                        }
                        //低密度脂蛋白
                        if (eachtest.testcode == "Y4.0040")
                        {
                            myWorkSheet.Cells[peoplecount, 86] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 87] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 88] = eachtest.testhigher;
                        }
                        //白细胞计数(WBC)
                        if (eachtest.testcode == "Y7.0100")
                        {
                            myWorkSheet.Cells[peoplecount, 90] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 91] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 92] = eachtest.testhigher;
                        }
                        //淋巴细胞绝对值
                        if (eachtest.testcode == "Y7.0120")
                        {
                            myWorkSheet.Cells[peoplecount, 93] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 94] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 95] = eachtest.testhigher;
                        }
                        //中间细胞绝对值
                        if (eachtest.testcode == "Y7.0380")
                        {
                            myWorkSheet.Cells[peoplecount, 96] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 97] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 98] = eachtest.testhigher;
                        }
                        //粒细胞绝对值
                        if (eachtest.testcode == "Y7.0134")
                        {
                            myWorkSheet.Cells[peoplecount, 99] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 100] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 101] = eachtest.testhigher;
                        }
                        //红细胞计数(RBC)
                        if (eachtest.testcode == "Y7.0010")
                        {
                            myWorkSheet.Cells[peoplecount, 102] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 103] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 104] = eachtest.testhigher;
                        }
                        //血红蛋白(HGB)
                        if (eachtest.testcode == "Y7.0020")
                        {
                            myWorkSheet.Cells[peoplecount, 105] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 106] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 107] = eachtest.testhigher;
                        }
                        //RBC平均HGB浓度(MCHC)
                        if (eachtest.testcode == "Y7.0060")
                        {
                            myWorkSheet.Cells[peoplecount, 108] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 109] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 110] = eachtest.testhigher;
                        }
                        //红细胞平均体积(MCV)
                        if (eachtest.testcode == "Y7.0040")
                        {
                            myWorkSheet.Cells[peoplecount, 111] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 112] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 113] = eachtest.testhigher;
                        }
                        //RBC平均HGB含量(MCH)
                        if (eachtest.testcode == "Y7.0400")
                        {
                            myWorkSheet.Cells[peoplecount, 114] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 115] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 116] = eachtest.testhigher;
                        }
                        //红细胞分布宽度（RDW）
                        if (eachtest.testcode == "Y7.0070")
                        {
                            myWorkSheet.Cells[peoplecount, 117] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 118] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 119] = eachtest.testhigher;
                        }
                        //红细胞压积(HCT) 又称红细胞比容
                        if (eachtest.testcode == "Y7.0275")
                        {
                            myWorkSheet.Cells[peoplecount, 120] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 121] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 122] = eachtest.testhigher;
                        }
                        //血小板计数（PLT）	
                        if (eachtest.testcode == "Y7.0210")
                        {
                            myWorkSheet.Cells[peoplecount, 123] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 124] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 125] = eachtest.testhigher;
                        }
                        //尿葡萄糖
                        if (eachtest.testcode == "Y8.0060")
                        {
                            myWorkSheet.Cells[peoplecount, 126] = eachtest.testresult;
                        }
                        //尿胆红素
                        if (eachtest.testcode == "Y8.0090")
                        {
                            myWorkSheet.Cells[peoplecount, 127] = eachtest.testresult;
                        }
                        //尿酮体
                        if (eachtest.testcode == "Y8.0070")
                        {
                            myWorkSheet.Cells[peoplecount, 128] = eachtest.testresult;
                        }
                        //尿比重
                        if (eachtest.testcode == "Y8.0040")
                        {
                            myWorkSheet.Cells[peoplecount, 129] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 130] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 131] = eachtest.testhigher;
                        }
                        //尿潜血、尿红细胞
                        if (eachtest.testcode == "Y8.0110")
                        {
                            myWorkSheet.Cells[peoplecount, 132] = eachtest.testresult;
                        }
                        //尿酸碱度
                        if (eachtest.testcode == "Y8.0030")
                        {
                            myWorkSheet.Cells[peoplecount, 133] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 134] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 135] = eachtest.testhigher;
                        }
                        //尿蛋白
                        if (eachtest.testcode == "Y8.0050")
                        {
                            myWorkSheet.Cells[peoplecount, 136] = eachtest.testresult;
                        }
                        //尿胆原
                        if (eachtest.testcode == "Y8.0080")
                        {
                            myWorkSheet.Cells[peoplecount, 137] = eachtest.testresult;
                        }
                        //尿亚硝酸盐
                        if (eachtest.testcode == "Y8.0100")
                        {
                            myWorkSheet.Cells[peoplecount, 138] = eachtest.testresult;
                        }
                        //尿白细胞酯酶
                        if (eachtest.testcode == "Y8.0120")
                        {
                            myWorkSheet.Cells[peoplecount, 139] = eachtest.testresult;
                        }
                        //大便颜色
                        if (eachtest.testcode == "Y9.0020")
                        {
                            myWorkSheet.Cells[peoplecount, 140] = eachtest.testresult;
                        }
                        //大便粘液 无检测项目

                        //大便潜血 免疫法 
                        if (eachtest.testcode == "Y9.0260")
                        {
                            myWorkSheet.Cells[peoplecount, 142] = eachtest.testresult;
                        }
                        //甲胎蛋白定量（AFP-N）		
                        if (eachtest.testcode == "YC.0030")
                        {
                            myWorkSheet.Cells[peoplecount, 143] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 144] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 145] = eachtest.testhigher;
                        }
                        //癌胚抗原定量（CEA-N）		
                        if (eachtest.testcode == "YC.0040")
                        {
                            myWorkSheet.Cells[peoplecount, 146] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 147] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 148] = eachtest.testhigher;
                        }
                        //糖类抗原（CA242）		
                        if (eachtest.testcode == "YC.0301")
                        {
                            myWorkSheet.Cells[peoplecount, 149] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 150] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 151] = eachtest.testhigher;
                        }
                        //男：前列腺特异抗原（PSA）		
                        if (eachtest.testcode == "YC.0120")
                        {
                            myWorkSheet.Cells[peoplecount, 152] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 153] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 154] = eachtest.testhigher;
                        }
                        //女：糖类抗原（CA15-3）		
                        if (eachtest.testcode == "YC.0160")
                        {
                            myWorkSheet.Cells[peoplecount, 155] = eachtest.testresult;
                            myWorkSheet.Cells[peoplecount, 156] = eachtest.testlower;
                            myWorkSheet.Cells[peoplecount, 157] = eachtest.testhigher;
                        }
                    }
                }
                finally{}
                //各个科室检查结果
                try
                {
                    var departmentResult = from s4 in myMedBaseEntities.hdatadep where checkpatient.membcode == s4.membcode && s4.checkdate > startDate && s4.checkdate < endDate select s4;
                    foreach (var eachdepart in departmentResult)
                    {
                        //内科
                        if (eachdepart.deptcode == "D00" || eachdepart.deptcode == "D01")
                        {
                            myWorkSheet.Cells[peoplecount, 26] = eachdepart.depresult.ToString();
                        }
                        //外科
                        if (eachdepart.deptcode == "E00") myWorkSheet.Cells[peoplecount, 28] = eachdepart.depresult.ToString();
                        //生殖泌尿
                        if (eachdepart.deptcode == "E01") myWorkSheet.Cells[peoplecount, 35] = eachdepart.depresult.ToString();
                        //耳鼻喉
                        if (eachdepart.deptcode == "H00" || eachdepart.deptcode == "H01") myWorkSheet.Cells[peoplecount, 36] = eachdepart.depresult.ToString();
                        //口腔
                        if (eachdepart.deptcode == "I00") myWorkSheet.Cells[peoplecount, 37] = eachdepart.depresult.ToString();
                        //眼科
                        if (eachdepart.deptcode == "G00") myWorkSheet.Cells[peoplecount, 38] = eachdepart.depresult.ToString();
                        //妇科
                        if (eachdepart.deptcode == "J00") myWorkSheet.Cells[peoplecount, 39] = eachdepart.depresult.ToString();
                        //神经科
                        if (eachdepart.deptcode == "F00") myWorkSheet.Cells[peoplecount, 40] = eachdepart.depresult.ToString();
                        //心电图
                        if (eachdepart.deptcode == "M00") myWorkSheet.Cells[peoplecount, 41] = eachdepart.depresult.ToString();
                        //胸片
                        if (eachdepart.deptcode == "K00") myWorkSheet.Cells[peoplecount, 42] = eachdepart.depresult.ToString();
                        //B超
                        if (eachdepart.deptcode == "L00") myWorkSheet.Cells[peoplecount, 43] = eachdepart.depresult.ToString();
                    }
                }
                catch
                {
                    //内科
                    myWorkSheet.Cells[peoplecount, 26] = "空白";
                    //外科
                    myWorkSheet.Cells[peoplecount, 28] = "空白";
                    //生殖泌尿
                    myWorkSheet.Cells[peoplecount, 35] = "空白";
                    //耳鼻喉
                    myWorkSheet.Cells[peoplecount, 36] = "空白";
                    //口腔
                    myWorkSheet.Cells[peoplecount, 37] = "空白";
                    //眼科
                    myWorkSheet.Cells[peoplecount, 38] = "空白";
                    //妇科
                    myWorkSheet.Cells[peoplecount, 39] = "空白";
                    //神经科
                    myWorkSheet.Cells[peoplecount, 40] = "空白";
                    //心电图
                    myWorkSheet.Cells[peoplecount, 41] = "空白";
                    //胸片
                    myWorkSheet.Cells[peoplecount, 42] = "空白";
                    //B超
                    myWorkSheet.Cells[peoplecount, 43] = "空白";
                }
                finally { }
                //10条主要疾病诊断
                try
                {
                    var diseaseResult = from s5 in myMedBaseEntities.hdatadiag where checkpatient.checkcode == s5.checkcode select s5;
                    int i = 0;
                    foreach (var eachDisease in diseaseResult)
                    {
                        myWorkSheet.Cells[peoplecount, 158 + i] = eachDisease.diagname;
                        myWorkSheet.Cells[peoplecount, 159 + i] = eachDisease.diagcode;
                        i=i+2;
                        if (i > 32) break;
                    }
                }
                catch
                {

                }
                finally { }
                //总检结论，总检建议
                try
                {
                    var totalResult = (from s6 in myMedBaseEntities.hdatarep where checkpatient.checkcode == s6.checkcode select s6).Single();
                    myWorkSheet.Cells[peoplecount, 190] = totalResult.hresult;
                    myWorkSheet.Cells[peoplecount, 191] = totalResult.hadvice;
                }
                catch { }
                finally { }
            }
            myWorkbook.SaveAs(FilePath);
            myWorkbook.Close();
            myExcel.Quit();
            iffinished.Text = "已完成！";
        }

        private void btn_selectSavePath_Click(object sender, EventArgs e)
        {
            OpenFileDialog myFileDialog = new OpenFileDialog();
            myFileDialog.Filter = "Excel|*.xls";
            if (myFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtbox_FilePath.Text = myFileDialog.FileName;
            }
        }

    }
}
