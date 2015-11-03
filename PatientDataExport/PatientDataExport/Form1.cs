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

        private void btn_beginProgress_Click(object sender, EventArgs e)
        {
            //禁用界面上面所有按钮
            btn_beginProgress.Enabled = false;
            btn_selectSavePath.Enabled = false;
            datePicker_startDate.Enabled = false;
            //患者的编号
            int peoplecount = 0;
            //设置要查询的时间
            DateTime startDate ;
            startDate = datePicker_startDate.Value;
            //控制输入的日期的有效值
            DateTime endDate;
            endDate = datePicker_startDate.Value.AddYears(1);
            lab_endDate.Text = endDate.ToString();
            //文件的存储路径
            String FilePath = txtbox_FilePath.Text;
            Excel.Application myExcel = new Excel.Application();
            myExcel.Visible = false;
            Excel.Workbook myWorkbook = myExcel.Workbooks.Add(true);
            Excel.Worksheet myWorkSheet = myWorkbook.Worksheets[1];


            medbase201507Entities1 myMedBaseEntities = new medbase201507Entities1();
            //查询所有的待查询时间段内检查的患者
            //查询条件  a0704 任职级别 01 副市级 02 正局级 03 副局级 04 正高 05 副高 14 院士
            //查询条件  a6405 在职情况 02 离休
            //&& (s1.a0704 == "01" || s1.a0704 == "02" || s1.a0704 == "03" || s1.a0704 == "04" || s1.a0704 == "05" || s1.a0704 == "14")
            var ExportResult = from s1 in myMedBaseEntities.hcheckmemb
                               where s1.checkdate > startDate && s1.checkdate < endDate && (s1.a0704 == "01" || s1.a0704 == "02" || s1.a0704 == "03" || s1.a0704 == "04" || s1.a0704 == "05" || s1.a0704 == "14" || s1.a6405 == "02")  
                               select s1;
            //总数
            totalNum.Text = ExportResult.Count().ToString();
            //遍历所有的患者
            foreach (var checkpatient in ExportResult)
            {
                //遍历每一位患者
                peoplecount++;
                progressNum.Text = peoplecount.ToString();
                //设置行高
                ((Excel.Range)myWorkSheet.Rows[peoplecount , System.Type.Missing]).RowHeight = 20;
                //患者的编号
                myWorkSheet.Cells[peoplecount, 1] = peoplecount;
                //姓名 a0101 使用membcode代替
                myWorkSheet.Cells[peoplecount, 2] = checkpatient.a0101;
                //性别
                myWorkSheet.Cells[peoplecount, 3] = checkpatient.a0107;
                //保健类型
                if (checkpatient.a0704 != null) 
                {
                    switch (checkpatient.a0704)
                    {
                        case "01": 
                            myWorkSheet.Cells[peoplecount, 8] = "1";
                            break;
                        case "02":
                            myWorkSheet.Cells[peoplecount, 8] = "1";
                            break;
                        case "03":
                            myWorkSheet.Cells[peoplecount, 8] = "1";
                            break;
                        case "04":
                            myWorkSheet.Cells[peoplecount, 8] = "3";
                            break;
                        case "05":
                            myWorkSheet.Cells[peoplecount, 8] = "3";
                            break;
                        case "14":
                            myWorkSheet.Cells[peoplecount, 8] = "3";
                            break;
                        default:
                            myWorkSheet.Cells[peoplecount, 8] = "无相符项目";
                            break;

                    }
                }
                //特殊的离休类型的处理
                if (checkpatient.a6405 == "02")
                {
                    myWorkSheet.Cells[peoplecount, 8] = "2";
                }
                //出生年月
                try
                {
                    //从basemember当中搜索必要的信息，如出生年月和身份证号
                    var searchBaseMemb= (from s2 in myMedBaseEntities.hbasememb where s2.membcode == checkpatient.membcode select s2).FirstOrDefault();
                    //出生年月
                    if (searchBaseMemb.a0111 == null)
                    {
                        try
                        {
                            int patientAge = 0;
                            patientAge = Convert.ToInt16(checkpatient.age);
                            myWorkSheet.Cells[peoplecount, 4] = startDate.AddYears( -patientAge);
                        }
                        catch
                        {
                            myWorkSheet.Cells[peoplecount, 4] = "空白";
                        }
                        
                    }
                    else
                    {
                        myWorkSheet.Cells[peoplecount, 4] = searchBaseMemb.a0111.ToString();
                    }
                    //移动电话
                    myWorkSheet.Cells[peoplecount, 6] = searchBaseMemb.mobileno.ToString();
                    //身份证号
                    myWorkSheet.Cells[peoplecount, 7] = searchBaseMemb.a0177.ToString();
                    //保健证号
                    myWorkSheet.Cells[peoplecount, 9] = searchBaseMemb.healthcode.ToString();
                }
                catch 
                {
                }
                finally { }
                //工作单位
                myWorkSheet.Cells[peoplecount, 5] = checkpatient.b0105.ToString();
                //体检医院
                myWorkSheet.Cells[peoplecount, 10] = "天津医科大学总医院";
                //各个检查结果
                try
                {
                    var testResult = from s3 in myMedBaseEntities.hdatadeptest where checkpatient.checkcode == s3.checkcode  select s3;
                    if (testResult == null)
                    { 
                        myWorkSheet.Cells[peoplecount, 194] = "没有相关的检验检查结果"; 
                    }
                    else
                    {
                        int testNum = 0;
                        foreach (var eachtest in testResult)
                        {
                            //输出所有的检查结果
                            testNum = testNum = 2;
                            //身高
                            if (eachtest.testcode == "D4.0010")
                            {
                                myWorkSheet.Cells[peoplecount, 23] = eachtest.testresult;
                            }
                            //体重
                            if (eachtest.testcode == "D4.0020")
                            {
                                myWorkSheet.Cells[peoplecount, 24] = eachtest.testresult;
                            }
                            //血压
                            if (eachtest.testcode == "D4.0040")
                            {
                                try
                                {
                                    string[] splitestring = eachtest.testresult.Split('/');
                                    myWorkSheet.Cells[peoplecount, 25] = splitestring[0].ToString();
                                    myWorkSheet.Cells[peoplecount, 26] = splitestring[1].ToString();
                                }
                                catch
                                {
                                    myWorkSheet.Cells[peoplecount, 25] = "空白";
                                    myWorkSheet.Cells[peoplecount, 26] = "空白";
                                }
                                finally
                                {
                                }
                            }
                            //谷丙转氨酶（ALT）
                            if (eachtest.testcode == "Y1.0010")
                            {
                                myWorkSheet.Cells[peoplecount, 46] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 47] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 48] = eachtest.testhigher;
                            }
                            //总胆红素(TBIL)
                            if (eachtest.testcode == "Y1.0050")
                            {
                                myWorkSheet.Cells[peoplecount, 49] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 50] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 51] = eachtest.testhigher;
                            }
                            //直接胆红素(DBIL)
                            if (eachtest.testcode == "Y1.0060")
                            {
                                myWorkSheet.Cells[peoplecount, 52] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 53] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 54] = eachtest.testhigher;
                            }
                            //总蛋白(TP)
                            if (eachtest.testcode == "Y1.0080")
                            {
                                myWorkSheet.Cells[peoplecount, 55] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 56] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 57] = eachtest.testhigher;
                            }
                            //白蛋白(ALB)
                            if (eachtest.testcode == "Y1.0090")
                            {
                                myWorkSheet.Cells[peoplecount, 58] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 59] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 60] = eachtest.testhigher;
                            }
                            //球蛋白(GLB)
                            if (eachtest.testcode == "Y1.0100")
                            {
                                myWorkSheet.Cells[peoplecount, 61] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 62] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 63] = eachtest.testhigher;
                            }
                            //尿素氮(BUN)
                            if (eachtest.testcode == "Y2.0010")
                            {
                                myWorkSheet.Cells[peoplecount, 64] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 65] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 66] = eachtest.testhigher;
                            }
                            //肌酐(Cr)
                            if (eachtest.testcode == "Y2.0020")
                            {
                                myWorkSheet.Cells[peoplecount, 67] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 68] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 69] = eachtest.testhigher;
                            }
                            //尿酸(UA)
                            if (eachtest.testcode == "Y2.0030")
                            {
                                myWorkSheet.Cells[peoplecount, 70] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 71] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 72] = eachtest.testhigher;
                            }
                            //葡萄糖(GLU)
                            if (eachtest.testcode == "Y3.0010")
                            {
                                myWorkSheet.Cells[peoplecount, 73] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 74] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 75] = eachtest.testhigher;
                            }
                            //糖化血红蛋白
                            if (eachtest.testcode == "Y3.0030")
                            {
                                myWorkSheet.Cells[peoplecount, 76] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 77] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 78] = eachtest.testhigher;
                            }
                            //总胆固醇(CHOL)
                            if (eachtest.testcode == "Y4.0010")
                            {
                                myWorkSheet.Cells[peoplecount, 79] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 80] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 81] = eachtest.testhigher;
                            }
                            //甘油三酯(TG)
                            if (eachtest.testcode == "Y4.0020")
                            {
                                myWorkSheet.Cells[peoplecount, 82] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 83] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 84] = eachtest.testhigher;
                            }
                            //高密度脂蛋白
                            if (eachtest.testcode == "Y4.0030")
                            {
                                myWorkSheet.Cells[peoplecount, 85] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 86] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 87] = eachtest.testhigher;
                            }
                            //低密度脂蛋白
                            if (eachtest.testcode == "Y4.0040")
                            {
                                myWorkSheet.Cells[peoplecount, 88] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 89] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 90] = eachtest.testhigher;
                            }
                            //白细胞计数(WBC)
                            if (eachtest.testcode == "Y7.0100")
                            {
                                myWorkSheet.Cells[peoplecount, 91] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 92] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 93] = eachtest.testhigher;
                            }
                            //淋巴细胞绝对值
                            if (eachtest.testcode == "Y7.0120")
                            {
                                myWorkSheet.Cells[peoplecount, 94] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 95] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 96] = eachtest.testhigher;
                            }
                            //中间细胞绝对值
                            if (eachtest.testcode == "Y7.0380")
                            {
                                myWorkSheet.Cells[peoplecount, 97] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 98] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 99] = eachtest.testhigher;
                            }
                            //粒细胞绝对值
                            if (eachtest.testcode == "Y7.0134")
                            {
                                myWorkSheet.Cells[peoplecount, 100] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 101] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 102] = eachtest.testhigher;
                            }
                            //红细胞计数(RBC)
                            if (eachtest.testcode == "Y7.0010")
                            {
                                myWorkSheet.Cells[peoplecount, 103] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 104] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 105] = eachtest.testhigher;
                            }
                            //血红蛋白(HGB)
                            if (eachtest.testcode == "Y7.0020")
                            {
                                myWorkSheet.Cells[peoplecount, 106] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 107] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 108] = eachtest.testhigher;
                            }
                            //RBC平均HGB浓度(MCHC)
                            if (eachtest.testcode == "Y7.0060")
                            {
                                myWorkSheet.Cells[peoplecount, 109] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 110] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 111] = eachtest.testhigher;
                            }
                            //红细胞平均体积(MCV)
                            if (eachtest.testcode == "Y7.0040")
                            {
                                myWorkSheet.Cells[peoplecount, 112] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 113] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 114] = eachtest.testhigher;
                            }
                            //RBC平均HGB含量(MCH)
                            if (eachtest.testcode == "Y7.0400")
                            {
                                myWorkSheet.Cells[peoplecount, 115] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 116] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 117] = eachtest.testhigher;
                            }
                            //红细胞分布宽度（RDW）
                            if (eachtest.testcode == "Y7.0070")
                            {
                                myWorkSheet.Cells[peoplecount, 118] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 119] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 120] = eachtest.testhigher;
                            }
                            //红细胞压积(HCT) 又称红细胞比容
                            if (eachtest.testcode == "Y7.0275")
                            {
                                myWorkSheet.Cells[peoplecount, 121] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 122] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 123] = eachtest.testhigher;
                            }
                            //血小板计数（PLT）	
                            if (eachtest.testcode == "Y7.0210")
                            {
                                myWorkSheet.Cells[peoplecount, 124] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 125] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 126] = eachtest.testhigher;
                            }
                            //尿葡萄糖
                            if (eachtest.testcode == "Y8.0060")
                            {
                                myWorkSheet.Cells[peoplecount, 127] = eachtest.testresult;
                            }
                            //尿胆红素
                            if (eachtest.testcode == "Y8.0090")
                            {
                                myWorkSheet.Cells[peoplecount, 128] = eachtest.testresult;
                            }
                            //尿酮体
                            if (eachtest.testcode == "Y8.0070")
                            {
                                myWorkSheet.Cells[peoplecount, 129] = eachtest.testresult;
                            }
                            //尿比重
                            if (eachtest.testcode == "Y8.0040")
                            {
                                myWorkSheet.Cells[peoplecount, 130] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 131] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 132] = eachtest.testhigher;
                            }
                            //尿潜血、尿红细胞
                            if (eachtest.testcode == "Y8.0110")
                            {
                                myWorkSheet.Cells[peoplecount, 133] = eachtest.testresult;
                            }
                            //尿酸碱度
                            if (eachtest.testcode == "Y8.0030")
                            {
                                myWorkSheet.Cells[peoplecount, 134] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 135] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 136] = eachtest.testhigher;
                            }
                            //尿蛋白
                            if (eachtest.testcode == "Y8.0050")
                            {
                                myWorkSheet.Cells[peoplecount, 137] = eachtest.testresult;
                            }
                            //尿胆原
                            if (eachtest.testcode == "Y8.0080")
                            {
                                myWorkSheet.Cells[peoplecount, 138] = eachtest.testresult;
                            }
                            //尿亚硝酸盐
                            if (eachtest.testcode == "Y8.0100")
                            {
                                myWorkSheet.Cells[peoplecount, 139] = eachtest.testresult;
                            }
                            //尿白细胞酯酶
                            if (eachtest.testcode == "Y8.0120")
                            {
                                myWorkSheet.Cells[peoplecount, 140] = eachtest.testresult;
                            }
                            //大便颜色
                            if (eachtest.testcode == "Y9.0020")
                            {
                                myWorkSheet.Cells[peoplecount, 141] = eachtest.testresult;
                            }
                            //大便粘液 无检测项目 142

                            //大便潜血 免疫法 
                            if (eachtest.testcode == "Y9.0260")
                            {
                                myWorkSheet.Cells[peoplecount, 143] = eachtest.testresult;
                            }
                            //甲胎蛋白定量（AFP-N）		
                            if (eachtest.testcode == "YC.0030")
                            {
                                myWorkSheet.Cells[peoplecount, 144] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 145] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 146] = eachtest.testhigher;
                            }
                            //癌胚抗原定量（CEA-N）		
                            if (eachtest.testcode == "YC.0040")
                            {
                                myWorkSheet.Cells[peoplecount, 147] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 148] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 149] = eachtest.testhigher;
                            }
                            //糖类抗原（CA242）		
                            if (eachtest.testcode == "YC.0301")
                            {
                                myWorkSheet.Cells[peoplecount, 150] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 151] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 152] = eachtest.testhigher;
                            }
                            //男：前列腺特异抗原（PSA）		
                            if (eachtest.testcode == "YC.0120")
                            {
                                myWorkSheet.Cells[peoplecount, 153] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 154] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 155] = eachtest.testhigher;
                            }
                            //女：糖类抗原（CA15-3）		
                            if (eachtest.testcode == "YC.0160")
                            {
                                myWorkSheet.Cells[peoplecount, 156] = eachtest.testresult;
                                myWorkSheet.Cells[peoplecount, 157] = eachtest.testlower;
                                myWorkSheet.Cells[peoplecount, 158] = eachtest.testhigher;
                            }
                        }
                    }
                }
                finally{}
                //各个科室检查结果
                try
                {
                    var departmentResult = from s4 in myMedBaseEntities.hdatadep where checkpatient.checkcode == s4.checkcode select s4;
                    //int depNum = 0;
                    if (departmentResult == null)
                    {
                        myWorkSheet.Cells[peoplecount, 194] = "没有相关科室的检查结果";
                    }
                    else
                    {
                        foreach (var eachdepart in departmentResult)
                        {
                            //输出所有的检查结果
                            //depNum = depNum + 2;
                            //内科
                            if (eachdepart.deptcode == "D00" || eachdepart.deptcode == "D01")
                            {
                                myWorkSheet.Cells[peoplecount, 27] = eachdepart.depresult.ToString();
                            }
                            //外科
                            if (eachdepart.deptcode == "E00") myWorkSheet.Cells[peoplecount, 29] = eachdepart.depresult.ToString();
                            //生殖泌尿
                            if (eachdepart.deptcode == "E01") myWorkSheet.Cells[peoplecount, 36] = eachdepart.depresult.ToString();
                            //耳鼻喉
                            if (eachdepart.deptcode == "H00" || eachdepart.deptcode == "H01") myWorkSheet.Cells[peoplecount, 36] = eachdepart.depresult.ToString();
                            //口腔
                            if (eachdepart.deptcode == "I00") myWorkSheet.Cells[peoplecount, 38] = eachdepart.depresult.ToString();
                            //眼科
                            if (eachdepart.deptcode == "G00") myWorkSheet.Cells[peoplecount, 39] = eachdepart.depresult.ToString();
                            //妇科
                            if (eachdepart.deptcode == "J00") myWorkSheet.Cells[peoplecount, 40] = eachdepart.depresult.ToString();
                            //神经科
                            if (eachdepart.deptcode == "F00") myWorkSheet.Cells[peoplecount, 41] = eachdepart.depresult.ToString();
                            //心电图
                            if (eachdepart.deptcode == "M00") myWorkSheet.Cells[peoplecount, 42] = eachdepart.depresult.ToString();
                            //胸片
                            if (eachdepart.deptcode == "K00") myWorkSheet.Cells[peoplecount, 43] = eachdepart.depresult.ToString();
                            //B超
                            if (eachdepart.deptcode == "L00") myWorkSheet.Cells[peoplecount, 44] = eachdepart.depresult.ToString();
                        }
                        
                    }
                }
                catch
                {
                }
                finally { }
                //10条主要疾病诊断
                try
                {
                    var diseaseResult = from s5 in myMedBaseEntities.hdatadiag where checkpatient.checkcode == s5.checkcode select s5;
                    int i = 0;
                    foreach (var eachDisease in diseaseResult)
                    {
                        myWorkSheet.Cells[peoplecount, 159 + i] = eachDisease.diagname;
                        myWorkSheet.Cells[peoplecount, 160 + i] = eachDisease.diagcode;
                        i=i+2;
                        if (i > 28) break;
                    }
                }
                catch { }
                finally { }
                //总检结论，总检建议
                try
                {
                    var totalResult = (from s6 in myMedBaseEntities.hdatarep where checkpatient.checkcode == s6.checkcode select s6).Single();
                    myWorkSheet.Cells[peoplecount, 189] = totalResult.hresult;
                    myWorkSheet.Cells[peoplecount, 190] = totalResult.hadvice;
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
            myFileDialog.Filter = "Excel|*.xlsx";
            if (myFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtbox_FilePath.Text = myFileDialog.FileName;
            }
        }

        private void datePicker_startDate_ValueChanged(object sender, EventArgs e)
        {
            if (datePicker_startDate.Value > Convert.ToDateTime("2017-1-1 00:00:00")) datePicker_startDate.Value = Convert.ToDateTime("2017-1-1 00:00:00");
            if (datePicker_startDate.Value < Convert.ToDateTime("2008-1-1 00:00:00")) datePicker_startDate.Value = Convert.ToDateTime("2008-1-1 00:00:00");
            lab_endDate.Text = datePicker_startDate.Value.AddYears(1).ToShortDateString();
        }


    }
}
