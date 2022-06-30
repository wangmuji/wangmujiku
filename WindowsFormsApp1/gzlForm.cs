using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace WindowsFormsApp1
{
    public partial class gzlForm : Form
    {
        public gzlForm()
        {
            InitializeComponent();
        }

        public static void smethod(IRow row)//学校计算方法
        {
            string course_name = row.GetCell(8).ToString();

            if (course_name.Contains("）") == true)
            {
                course_name = course_name.Replace("）", "");
            }
            string[] arrays = course_name.Split('（');
            course_name = arrays[0];
            int num = 0;
            foreach (string i in arrays)
                num++;

            double H = Double.Parse(row.GetCell(7).ToString());//H为计划课时数
            double N = Double.Parse(row.GetCell(6).ToString());//学生数
            double K1 = 1;//K1为理论班合班系数
            if (N > 60)
            {
                K1 = 0.005 * N + 0.75;
            }
            double n = ((H / 6) > 8) ? 8 : (H / 6);//作业次数
            n = Math.Round(n);
            double Z = Math.Round(H / 16);//教学周数
            double M = ((N / 30) > 1) ? N / 30 : 1;//实际班数
            double K2 = 1;//实验课容量系数
            if (N >= 30)
                K2 = 1.35;
            if (N < 29 && N >= 20)
                K2 = 1.2;
            if (N >= 10 && n < 19)
                K2 = 1.05;
            if (N >= 5 && N < 9)
                K2 = 0.9;
            if (N <= 4)
                K2 = 0.75;
            double K3 = 1.0;//实习地点系数
            double assistance = 0;//助课量
            double lead = H * K1 + N * n / 50;//主讲量
            double work = N * n / 50;//作业
            double experiment_hard = Math.Round(H * (1 + (M - 1) * 0.4) * K2, 3);//实验
            double experiment_soft = Math.Round(H * (0.5 + (M - 1) * 0.5) * K2, 3);
            double experiment_class = Math.Round(12 / 15 * N * Z * K3, 3);
            switch (course_name)
            {
                case "理论课":
                    {
                        //         sum = Math.Round((H * K1 + N * n / 50), 2);
                        if (arrays[num - 1] == "外语教材")
                        {
                            row.GetCell(9).SetCellValue(lead * 1.5);
                            row.GetCell(12).SetCellValue(work);
                            row.GetCell(17).SetCellValue(Math.Round(lead + work));
                        }
                        else
                        {
                            if (arrays[num - 1] == "双语")
                            {
                                row.GetCell(9).SetCellValue(lead * 1.2);
                                row.GetCell(12).SetCellValue(work);
                                row.GetCell(17).SetCellValue(Math.Round(lead + work));
                            }
                            else
                            {
                                row.GetCell(9).SetCellValue(lead);
                                row.GetCell(12).SetCellValue(work);
                                row.GetCell(17).SetCellValue(Math.Round(lead + work));
                            }
                        }
                        break;
                    }
                case "理论课-有助课":
                    {
                        //  sum = Math.Round(H * K1);
                        if (arrays[num - 1] == "MOOC")
                        {
                            row.GetCell(9).SetCellValue((H * K1 + N * n / 50) * 1.5);
                            row.GetCell(17).SetCellValue(Math.Round((H * K1 + N * n / 50) * 1.5));
                        }
                        else
                        {
                            row.GetCell(9).SetCellValue((H * K1 + N * n / 50));
                            row.GetCell(17).SetCellValue(Math.Round((H * K1 + N * n / 50)));
                        }
                        break;
                    }
                case "助课":
                    {
                        //    sum = Math.Round(N * n / 50);
                        if (arrays[num - 1] == "MOOC")
                        {
                            row.GetCell(12).SetCellValue(N * n / 50);
                            row.GetCell(17).SetCellValue(Math.Round(N * n / 50 * 2.0));
                        }
                        else
                        {
                            row.GetCell(12).SetCellValue(N * n / 50);
                            row.GetCell(17).SetCellValue(Math.Round(N * n / 50));
                        }
                        break;
                    }
                case "硬件实验":
                    {
                        // sum = Math.Round((H * (1 + (M - 1) * 0.4) * K2), 2);
                        row.GetCell(10).SetCellValue(experiment_hard);
                        row.GetCell(17).SetCellValue(Math.Round(experiment_hard));
                        break;
                    }
                case "软件实验":
                    {
                        //      sum = Math.Round((H * (0.5 + (M - 1) * 0.5) * K2), 2);
                        row.GetCell(10).SetCellValue(experiment_soft);
                        row.GetCell(17).SetCellValue(Math.Round(experiment_soft));
                        break;
                    }
                case "课程设计":
                    {
                        //       sum = Math.Round((12 * (1 / 15) * N * Z * K3), 2);
                        row.GetCell(10).SetCellValue(experiment_class);
                        row.GetCell(17).SetCellValue(Math.Round(experiment_class));
                        break;
                    }
            }
            //   return sum;
        }
        public static void Ssamecourse(ISheet sheet)
        {
            int q = 3;
            IRow row = sheet.GetRow(q);
            int i = q + 1;
            while (q < sheet.LastRowNum - 1)
            {
                while (sheet.GetRow(i).GetCell(5).ToString() == row.GetCell(5).ToString())
                {
                    if ((sheet.GetRow(i).GetCell(1).ToString() == row.GetCell(1).ToString()))
                    {
                        String name = sheet.GetRow(i).GetCell(5).ToString();
                        if (name.Contains("）") == true)
                        {
                            name = name.Replace("）", "");
                        }
                        string[] arrays = name.Split('（');
                        name = arrays[0];
                        int num = 0;
                        foreach (string s in arrays)
                            num++;
                        if (name == "软件实验" || name == "硬件实验" || name == "助课" || name == "课程设计")
                            break;
                        double H = Double.Parse(row.GetCell(7).ToString());//H为计划课时数
                        double N = Double.Parse(row.GetCell(6).ToString());//学生数
                        double K1 = 1;//K1为理论班合班系数
                        if (N > 60)
                        {
                            K1 = 0.005 * N + 0.75;
                        }
                        double n = ((H / 6) > 8) ? 8 : (H / 6);//作业次数
                        double lead = H * K1 + N * n / 50;//主讲量
                        switch (name)
                        {
                            case "理论课":
                                {
                                    if (arrays[num - 1] == "外语教材")
                                    {
                                        row.GetCell(9).SetCellValue(lead * 1.5 * 0.9);

                                    }
                                    else
                                    {
                                        if (arrays[num - 1] == "双语")
                                        {
                                            row.GetCell(9).SetCellValue(lead * 1.2 * 0.9);

                                        }
                                        else
                                        {
                                            row.GetCell(9).SetCellValue(lead * 0.9);

                                        }
                                    }
                                    break;
                                }
                            case "理论课-有助课":
                                {
                                    if (arrays[num - 1] == "MOOC")
                                    {
                                        row.GetCell(9).SetCellValue((H * K1 + N * n / 50) * 1.5 * 0.9);

                                    }
                                    else
                                    {
                                        row.GetCell(9).SetCellValue((H * K1 + N * n / 50) * 0.9);

                                    }
                                    break;
                                }
                        }
                    }
                }
                q++;
                i = q + 1;
                if (q >= sheet.LastRowNum)
                    break;
            }

        }
        private void button2_Click(object sender, EventArgs e)
        {
            string rfilepath = "";
            OpenFileDialog dialog = new OpenFileDialog();
            //dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择需处理教务文件";
            dialog.Filter = "所有文件(*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                rfilepath = dialog.FileName;
                IWorkbook wk = null;
                string extension = System.IO.Path.GetExtension(rfilepath);
                try
                {
                    FileStream fs = File.OpenRead(rfilepath);
                    if (extension.Equals(".xls"))
                    {
                        //把xls文件中的数据写入wk中
                        wk = new HSSFWorkbook(fs);
                    }
                    else
                    {
                        //把xlsx文件中的数据写入wk中
                        wk = new XSSFWorkbook(fs);
                    }

                    fs.Close();
                    //读取当前表数据
                    ISheet sheet = wk.GetSheetAt(0);
                    IRow row = sheet.GetRow(0);
                    int col = row.LastCellNum;
                    //LastRowNum 是当前表的总行数-1（注意）
                    int offset = 0;
                    int table = 2;
                    string s = row.GetCell(0).ToString();
                    if (s.Contains("校") && s.Contains("工作量"))
                    {
                        table = 3;
                    }
                    else
                    {
                        MessageBox.Show("失败");
                        return;
                    }
                    for (int i = table; i <= sheet.LastRowNum; i++)
                    {
                        row = sheet.GetRow(i);  //读取当前行数据

                        //   Console.WriteLine(i);
                        if (row != null && row.GetCell(0) != null)
                        {
                            smethod(row);

                        }
                    }
                    Ssamecourse(sheet);
                    FileStream file = new FileStream(rfilepath, FileMode.Open, FileAccess.Write);
                    wk.Write(file);
                    file.Close();
                    wk.Close();
                    MessageBox.Show("成功");
                }
                catch (Exception ex)
                {
                    //只在Debug模式下才输出
                    Console.WriteLine(ex.Message);
                }

            }
        }
        public static void Cmethod(IRow row)//学院计算方法
        {
            string course_name = row.GetCell(6).ToString();
            if (course_name.Contains(")"))
            {
                course_name.Replace(")", "");
            }
            string[] arrays = course_name.Split('(');
            course_name = arrays[0];
            double H = Math.Round(Double.Parse(row.GetCell(4).ToString()));//计划课时数
            double N = Math.Round(Double.Parse(row.GetCell(3).ToString()));//学生数
            double M = Math.Round(Double.Parse(row.GetCell(3).ToString()) / 30, 2);//实际班数
            if (M < 1)
                M = 1;
            double a = Math.Round(H * (1 + 0.1 * (M - 1)), 2);      //主讲学时
            double homework = Math.Round((N * H / 200), 2);     //作业&&非主讲学时
            double p = H * (1 + 0.1 * (M - 1)) * 0.4;       //实验主讲学时
            double P = H * 0.6;                 //实验非主讲学时
            double PP = H * (1 + 0.1 * (M - 1)) * 0.2;      //课设主讲学时
            double PS = H * 0.8;//课设非主讲学时
                                //    double sum = 0;//总工作量
                                //Console.WriteLine(a);
            switch (course_name)
            {
                case "理论课":
                    {
                        //      sum = Math.Round((N * H / 200) + H * (1 + 0.1 * (M - 1)),2);
                        row.GetCell(5).SetCellValue(M);
                        row.GetCell(12).SetCellValue(a);
                        row.GetCell(13).SetCellValue(homework);
                        row.GetCell(15).SetCellValue(homework);
                        row.GetCell(16).SetCellValue(H);
                        break;
                    }
                case "理论课-有助课":
                    {
                        //       sum = Math.Round(H * (1 + 0.1 * (M - 1)),2);
                        row.GetCell(12).SetCellValue(a);
                        row.GetCell(15).SetCellValue(0);
                        row.GetCell(16).SetCellValue(H);
                        break;
                    }
                case "助课":
                    {
                        //       sum = Math.Round(N * H / 200,2);
                        row.GetCell(13).SetCellValue(homework);
                        row.GetCell(14).SetCellValue(homework);
                        break;
                    }
                case "硬件实验":
                    {
                        //       sum = Math.Round(H*(1+0.1*(M-1)*0.4)+H*0.6,2);
                        row.GetCell(12).SetCellValue(p);
                        row.GetCell(14).SetCellValue(P);
                        row.GetCell(15).SetCellValue(P);
                        break;
                    }
                case "软件实验":
                    {
                        //        sum = Math.Round(H, 2);
                        row.GetCell(14).SetCellValue(H);
                        row.GetCell(15).SetCellValue(H);
                        break;
                    }
                case "课程设计":
                    {
                        //        sum = Math.Round(H*(1+0.1*(M-1)*0.2), 2);
                        row.GetCell(12).SetCellValue(PP);
                        row.GetCell(14).SetCellValue(PS);
                        row.GetCell(15).SetCellValue(PS);
                        break;
                    }
            }
            //  return sum;
        }
        public static void Csamecourse(ISheet sheet)//学院相同课程同一教师工作量计算
        {
            int q = 2;
            IRow row = sheet.GetRow(q);
            //       Console.WriteLine(row.GetCell(0).ToString());
            int i = q + 1;
            while (q < sheet.LastRowNum - 1 && row.GetCell(0) != null)
            {
                //           Console.WriteLine("进来了"+q+" "+i);
                while (sheet.GetRow(i).GetCell(2).ToString() == row.GetCell(2).ToString() && sheet.GetRow(i).GetCell(2) != null)
                {

                    if ((sheet.GetRow(i).GetCell(1).ToString() == row.GetCell(1).ToString()))
                    {
                        String name = sheet.GetRow(i).GetCell(6).ToString();
                        //        Console.WriteLine("这也进来了"+" "+name); 
                        if (name == "软件实验" || name == "助课")
                        {
                            break;
                        }
                        double H = Math.Round(Double.Parse(sheet.GetRow(i).GetCell(4).ToString()));//计划课时数
                        double N = Math.Round(Double.Parse(sheet.GetRow(i).GetCell(3).ToString()));//学生数
                        double M = Math.Round(Double.Parse(sheet.GetRow(i).GetCell(3).ToString()) / 30, 2);//实际班数
                        if (M < 1)
                            M = 1;
                        double a = Math.Round(H * (1 + 0.1 * (M - 1)), 2);      //主讲学时
                        double p = H * (1 + 0.1 * (M - 1)) * 0.4;       //实验主讲学时
                        double PP = H * (1 + 0.1 * (M - 1)) * 0.2;      //课设主讲学时
                        switch (name)
                        {
                            case "理论课":
                                {
                                    sheet.GetRow(i).GetCell(12).SetCellValue(0.9 * a);
                                    break;
                                }
                            case "硬件实验":
                                {
                                    sheet.GetRow(i).GetCell(12).SetCellValue(0.9 * p);
                                    break;
                                }
                            case "课程设计":
                                {
                                    sheet.GetRow(i).GetCell(12).SetCellValue(0.9 * PP);
                                    break;
                                }
                            case "理论课-有助课":
                                {
                                    sheet.GetRow(i).GetCell(12).SetCellValue(0.9 * a);
                                    break;
                                }
                        }
                    }
                    i++;
                    if (i >= sheet.LastRowNum)
                        break;
                }
                //        Console.WriteLine("进来了" + q + " " + i);
                q++;
                i = q + 1;
                if (q >= sheet.LastRowNum)
                    break;

            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            string rfilepath = "";
            OpenFileDialog dialog = new OpenFileDialog();
            //dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择需处理工作量文件";
            dialog.Filter = "所有文件(*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                rfilepath = dialog.FileName;
                IWorkbook wk = null;
                string extension = System.IO.Path.GetExtension(rfilepath);
                try
                {
                    FileStream fs = File.OpenRead(rfilepath);
                    if (extension.Equals(".xls"))
                    {
                        //把xls文件中的数据写入wk中
                        wk = new HSSFWorkbook(fs);
                    }
                    else
                    {
                        //把xlsx文件中的数据写入wk中
                        wk = new XSSFWorkbook(fs);
                    }

                    fs.Close();
                    //读取当前表数据
                    ISheet sheet = wk.GetSheetAt(0);
                    IRow row = sheet.GetRow(0);
                    int col = row.LastCellNum;
                    //LastRowNum 是当前表的总行数-1（注意）
                    int offset = 0;
                    int table = 2;
                    string s = row.GetCell(0).ToString();
                    if (s.Contains("学院") && s.Contains("工作量"))
                    {
                        table = 2;
                    }
                    else
                    {
                        MessageBox.Show("院工作量统计失败");
                        return;
                    }
                    for (int i = table; i <= sheet.LastRowNum; i++)
                    {
                        row = sheet.GetRow(i);  //读取当前行数据

                        //   Console.WriteLine(i);
                        if (row != null && row.GetCell(0) != null)
                        {
                            Cmethod(row);
                        }
                    }
                    Csamecourse(sheet);
                    FileStream file = new FileStream(rfilepath, FileMode.Open, FileAccess.Write);
                    wk.Write(file);
                    file.Close();
                    wk.Close();
                    MessageBox.Show("院工作量统计成功");
                }
                catch (Exception ex)
                {
                    //只在Debug模式下才输出
                    Console.WriteLine(ex.Message);
                }

            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
