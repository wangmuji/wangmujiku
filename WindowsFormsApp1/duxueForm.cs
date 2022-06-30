using System;

using NPOI;
using NPOI.HPSF;
using NPOI.SS.UserModel;
using System.Data;
using System.IO;
using NPOI.XSSF.UserModel;
using System.Linq;
using NPOI.HSSF.UserModel;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Text;
using System.Threading.Tasks;


namespace WindowsFormsApp1
{
    

    public partial class duxueForm : Form
    {
        public duxueForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int flag = 0;
            flag = duxue.AllWeekDuxue();
            if (flag == 1)
            {
                MessageBox.Show("导出成功");
            }
            else if (flag == 0)
            {
                MessageBox.Show("已取消导出");
            }
            else if (flag == -1)
            {
                MessageBox.Show("导出失败");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int flag=0;
            if (weekBox.Text == "")
            {
                MessageBox.Show("输入为空");

            }
            else
            {
                int whichWeek = int.Parse(weekBox.Text);
                if (whichWeek > 0 && whichWeek < 22)
                {
                    flag = duxue.OneWeekDuxue(whichWeek);
                }
                if (flag == 1)
                {
                    MessageBox.Show("导出成功");
                }
                else if (flag == 0)
                {
                    MessageBox.Show("已取消导出");
                }
                else if (flag == -1)
                {
                    MessageBox.Show("导出失败");
                }
            }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void weekBox_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void weekBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((int)e.KeyChar < 48 || (int)e.KeyChar > 57) && (int)e.KeyChar != 8 )
                e.Handled = true;


            if (weekBox.Text.Length >= 2)
            {
                if(Char.IsDigit(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
            
        }
    }


    public class Information
    {
        //        public int number = 0;
        public int day = 0;         //周几
        public string whichLesson = "Lesson";       //哪节课
        public string lessenName = "Name";          //课程名字  
        public string whichClass = "Class";         //哪些班
        public string whichGrade = "Grade";         //哪个年级
        public string whichWeek = "Week";           //哪周
        public string whichPlace = "Place";         //地点
        public string teacher = "teacher";          //教师
        public int personNum;                   //人数
        public Information next = null;         //下一节点

    }

    public class duxue
    {
        public static Information head = null;
        public static Information rear = null;
        public static int getDay(string time)   //解析周几上课
        {
            int day = 0;
            string dday;
            string pattern_day = "周[一, 二, 三, s四, 五, 六, 日]";
            dday = Regex.Match(time, pattern_day).ToString();
            if (dday.Equals("周一"))
            {
                day = 1;
            }
            else if (dday.Equals("周二"))
            {
                day = 2;
            }
            else if (dday.Equals("周三"))
            {
                day = 3;
            }
            else if (dday.Equals("周四"))
            {
                day = 4;
            }
            else if (dday.Equals("周五"))
            {
                day = 5;
            }
            else if (dday.Equals("周六"))
            {
                day = 6;
            }
            else if (dday.Equals("周日"))
            {
                day = 7;
            }
            return day;
        }
        public static string getLesson(string time)   //解析第几节上课
        {
            string Lesson;
            string pattern_lesson = "第.*节";
            MatchCollection LL;
            Lesson = Regex.Match(time, pattern_lesson).ToString();
            LL = Regex.Matches(Lesson, "[1-9]\\d*");
            Match aa = LL[0];
            string first = aa.ToString();
            aa = LL[LL.Count - 1];
            string last = aa.ToString();

            //Console.WriteLine("yes"+first+"--"+last);


            return first + "-" + last;
        }

        public static void getGradeAndClass(string value, Information inf)   //解析哪些班上课
        {
            int[] allGrade = new int[4];
            int grade = 0;
            string[] classInf = value.Split(',');       //以逗号分割
            string gg;
            string finalGrade = "";
            for (int i = 0; i < classInf.Length; i++)
            {
                gg = classInf[i];
                gg = gg.Substring(2, 2);        //截取中间两个数字,表示年级

                if (grade < int.Parse(gg))
                {
                    grade = int.Parse(gg);
                }

            }
            finalGrade += "20" + grade + "级";
            string c = "";
            int[] cla = new int[40];
            int first = 0;
            int last = 99;

            string sfirst;
            string slast;
            for (int i = 0; i < classInf.Length; i++)
            {
                gg = classInf[i];
                gg = gg.Substring(2, 4);        //截取后四位，代表年级班级，便于后续计算，例如1816
                cla[i] = int.Parse(gg);
                //               Console.WriteLine(cla[i]);
            }
            for (int i = 0; i < classInf.Length; i++)
            {
                if (cla[i] < grade * 100)    //排除高年级
                    continue;
                else if (first == 0)
                {
                    first = cla[i];
                }
                if (cla[i] + 1 != cla[i + 1] || i + 1 == classInf.Length)
                {
                    //                   Console.WriteLine(i + "  " + cla[i]);

                    last = cla[i];
                    sfirst = first.ToString().Substring(2, 2);      //截取后两个数字，代表班级
                    slast = last.ToString().Substring(2, 2);

                    if (first != last)
                    {
                        c += sfirst + "-" + slast;

                    }
                    else
                    {
                        if (first % 100 == 16)
                        {

                            c += "唐";
                        }
                        else
                        {
                            c += sfirst;        //单个班时不需要-
                        }
                    }
                    first = cla[i + 1];     //新的首

                    if (i + 1 != classInf.Length)
                    {
                        c += ",";
                    }

                }
            }
            inf.whichGrade = finalGrade;
            inf.whichClass = c;

            //Console.WriteLine("yes" + inf.whichGrade);
            //Console.WriteLine("no" + inf.whichClass+"\n\n");
            return;
        }
        public static string getWeek(string time)   //解析哪周上课
        {
            string week;
            string pattern_week = "{.*}";
            week = Regex.Match(time, pattern_week).ToString();
            week = week.Replace("{第", "");
            week = week.Replace("周}", "");
            return week;
        }
        public static void insert(Information inf)
        {
            Information ii = head;
            int jfirst;         //inf第几节上课
            int jlast;          //inf第几节下课
            int iifirst;
            int iilast;
            while (ii.next != null && inf.day >= ii.next.day)
            {

                if (inf.day == ii.next.day)
                {
                    iifirst = int.Parse(ii.next.whichLesson.Substring(0, 1));
                    iilast = int.Parse(ii.next.whichLesson.Substring(2, 1));
                    jfirst = int.Parse(inf.whichLesson.Substring(0, 1));
                    jlast = int.Parse(inf.whichLesson.Substring(2, 1));
                    if (jfirst < iifirst || (jfirst == iifirst && jlast < iilast))      //按第一节课上课时间排，先上课的排在前面
                    {
                        break;
                    }
                }
                ii = ii.next;
            }
            if (ii.next == null)       //说明在尾部了
            {
                ii.next = inf;
                rear = inf;
            }
            else
            {
                inf.next = ii.next;
                ii.next = inf;
            }
        }

        public static string getPlace(string p)
        {
            string shortPlace = "";
            if (p.Contains("经信"))
            {
                shortPlace += "经信";
            }
            else if (p.Contains("逸夫"))
            {
                shortPlace += "逸";
            }
            else if (p.Contains("计算机"))
            {
                shortPlace += "计算机楼";
            }
            else if (p.Contains("第三教学楼"))
            {
                shortPlace += "三教";
            }
            else if (p.Contains("李四光"))
            {
                shortPlace += "李四光";
            }
            else if (p.Contains("萃文"))
            {
                shortPlace += "萃文";
            }
            p = p.Split('#')[1];
            if (p.Contains("阶梯"))
            {
                p = p.Replace("第", "");
                p = p.Replace("阶梯", "");
                if (p.Equals("一")) p = "1阶";
                else if (p.Equals("二")) p = "2阶";
                else if (p.Equals("三")) p = "3阶";
                else if (p.Equals("四")) p = "4阶";
                else if (p.Equals("五")) p = "5阶";
                else if (p.Equals("六")) p = "6阶";
                else if (p.Equals("七")) p = "7阶";
                else if (p.Equals("八")) p = "8阶";
                else if (p.Equals("九")) p = "9阶";
                else if (p.Equals("十")) p = "10阶";
                else if (p.Equals("十一")) p = "11阶";
                else if (p.Equals("十二")) p = "12阶";
                else if (p.Equals("十三")) p = "13阶";
                else if (p.Equals("十四")) p = "14阶";
                else if (p.Equals("十五")) p = "15阶";
                else if (p.Equals("十六")) p = "16阶";
                else if (p.Equals("十七")) p = "17阶";
                else if (p.Equals("十八")) p = "18阶";
                else if (p.Equals("十九")) p = "19阶";
                else if (p.Equals("二十")) p = "20阶";
            }
            shortPlace += p;
            return shortPlace;
        }
        public static string switchDay(int day)
        {
            string d = "";
            if (day == 1) d = "一";
            if (day == 2) d = "二";
            if (day == 3) d = "三";
            if (day == 4) d = "四";
            if (day == 5) d = "五";
            if (day == 6) d = "六";
            if (day == 7) d = "日";

            return d;
        }

        public static int readJiaowu(string filePath)
        {

            IWorkbook wk = null;

            string extension = System.IO.Path.GetExtension(filePath);
            try
            {
                FileStream fs = File.OpenRead(filePath);
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
                //读取当前表数据。
                ISheet sheet = wk.GetSheetAt(0);

                IRow row = sheet.GetRow(0);  //读取当前行数据

                //LastRowNum 是当前表的总行数-1（注意）
                int offset = 0;
                int NUM = 1;
                string[] place;     //地点中间变量
                string[] time;      //时间中间变量
                int nn = 1;
                head = new Information();

                head.next = null;
                Information[] inf = new Information[10];

                row = sheet.GetRow(0);
                if(row!=null)
                {
                    if(row.GetCell(0)==null)
                    {
                        return -1;
                    }
                    else
                    {
                        if (!row.GetCell(0).ToString().Equals("学院开课信息列表"))          //判断文件依据
                        {
                            
                            return -1;
                        }
                    }
                }

                for (int i = 3; i <= sheet.LastRowNum; i++)
                {
                    row = sheet.GetRow(i);  //读取当前行数据
                    if (row != null)
                    {
                        inf = null;         //赋空
                        inf = new Information[10];          //
                        //LastCellNum 是当前行的总列数
                        for (int j = row.LastCellNum - 1; j >= 0; j--)        //从最后一列开始读，保证知道一周有几次课
                        {
                            //读取该行的第j列数据
                            string value=string.Empty;
                            if (row.GetCell(j) != null)
                                value = row.GetCell(j).ToString();
                            else
                                continue;       //该行的这列为空，多半文件出错或是选错了文件
                            place = null;

                            if (j == 7 && i > 2)    //上课地点,上课时间
                            {
                                place = value.Split(';');
                                NUM = value.Split(';').Length;      //获得该课每周上课次数，从而确定有几行
                                time = row.GetCell(j - 1).ToString().Split(';');      //每周的上课时间

                                for (int cc = 0; cc < NUM; cc++)
                                {
                                    inf[cc] = new Information();
                                    inf[cc].next = null;
                                    inf[cc].whichWeek = getWeek(time[cc]);
                                    inf[cc].whichLesson = getLesson(time[cc]);
                                    inf[cc].day = getDay(time[cc]);
                                    inf[cc].whichPlace = getPlace(place[cc]);
                                    //                                    inf[cc].number = nn;
                                    nn++;
                                    if (head.next == null)
                                    {
                                        head.next = inf[cc];
                                        rear = inf[cc];
                                    }
                                    else        //在此处插入链表
                                    {
                                        insert(inf[cc]);

                                    }
                                    //                                   Console.WriteLine(time[cc]);
                                }
                            }

                            if (j == 5)   //上课人数
                            {
                                for (int cc = 0; cc < NUM; cc++)
                                {
                                    inf[cc].personNum = int.Parse(value);

                                }

                            }
                            if (j == 4)    //班级,年级信息
                            {
                                for (int cc = 0; cc < NUM; cc++)
                                {
                                    //Console.WriteLine(value);
                                    getGradeAndClass(value, inf[cc]);
                                }
                                j--;            //j==3时的数据为学时，生成督学表时不需要
                            }

                            if (j == 2)    //课程名称
                            {
                                for (int cc = 0; cc < NUM; cc++)
                                {
                                    try
                                    {
                                        inf[cc].lessenName = value;
                                    }catch(Exception e)
                                    {
                                        MessageBox.Show("文件类型错误");
                                    }
                                }

                            }
                            if (j == 1)    //教师姓名
                            {
                                for (int cc = 0; cc < NUM; cc++)
                                {
                                    inf[cc].teacher = value;
                                }
                            }
                            //Console.Write(value.ToString() + " ");
                        }
                        //Console.WriteLine("\n");
                    }
                }

            }
            catch (Exception e)
            {
                //只在Debug模式下才输出
                return -1;
            }
            return 1;
        }

        public static int OneWeekDuxue(int week)     //根据输入第几周来生成督学表单
        {
            string rfilepath = "";


            OpenFileDialog dialog = new OpenFileDialog();
            //dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择需处理教务文件";
            dialog.Filter = "所有文件(*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                rfilepath = dialog.FileName;

            }
            else
                return 0;   //表示取消
            //string rfilepath = rpath;
            string wfilepath = "";          //写的路径

            if(readJiaowu(rfilepath)==-1)
            {
                MessageBox.Show("选错文件或所选教务信息文件已打开");
                return -1;
            }

            FolderBrowserDialog dialog1 = new FolderBrowserDialog();

            dialog1.Description = "请选择导出文件路径";

            if (dialog1.ShowDialog() == DialogResult.OK)
            {
                wfilepath = dialog1.SelectedPath + @"\";
            }
            else
                return 0;

            Information ii = head.next;
            int firstWeek;
            int lastWeek;
            int num = 2;
            int lieshu = 14;
            int realLieshu = 10;
            string[] headRow = { "1", "星期", "节", "课程名称", "地点", "推荐班级", "年级", "周", "任课教师", "人数", "教师情况", "学生情况（出勤率、迟到人数）", "教学效果", "意见和建议" };
            wfilepath += "/督学.xlsx";
            IWorkbook wb = null;
            try
            {
                FileStream fs = File.Create(wfilepath);
                wb = new HSSFWorkbook();
                ISheet sheet1 = wb.CreateSheet("sheet1");
                sheet1.CreateRow(0);
                for (int i = 0; i < lieshu; i++)       //创建表头
                {
                    sheet1.GetRow(0).CreateCell(i);
                    sheet1.GetRow(0).GetCell(i).SetCellValue(headRow[i]);

                }
                while (ii != null)
                {
                    firstWeek = int.Parse(ii.whichWeek.Split('-')[0]);
                    lastWeek = int.Parse(ii.whichWeek.Split('-')[1]);
                    if (firstWeek > week || lastWeek < week)           //说明此课程不在所选周中
                    {
                        ii = ii.next;
                        continue;               //重新循环
                    }
                    sheet1.CreateRow(num - 1);
                    for (int i = 0; i < realLieshu; i++)
                    {

                        sheet1.GetRow(num - 1).CreateCell(i);
                    }
                    sheet1.GetRow(num - 1).GetCell(0).SetCellValue(num.ToString());
                    sheet1.GetRow(num - 1).GetCell(1).SetCellValue(switchDay(ii.day));
                    sheet1.GetRow(num - 1).GetCell(2).SetCellValue(ii.whichLesson);
                    sheet1.GetRow(num - 1).GetCell(3).SetCellValue(ii.lessenName);
                    sheet1.GetRow(num - 1).GetCell(4).SetCellValue(ii.whichPlace);
                    sheet1.GetRow(num - 1).GetCell(5).SetCellValue(ii.whichClass);
                    sheet1.GetRow(num - 1).GetCell(6).SetCellValue(ii.whichGrade);
                    sheet1.GetRow(num - 1).GetCell(7).SetCellValue(ii.whichWeek);
                    sheet1.GetRow(num - 1).GetCell(8).SetCellValue(ii.teacher);
                    sheet1.GetRow(num - 1).GetCell(9).SetCellValue(ii.personNum);
                    num++;
                    ii = ii.next;
                }

                Console.WriteLine(num);
                wb.Write(fs);

                fs.Close();
                wb.Close();
            }
            catch (Exception e)
            {
                return -1;//表示失败
            }
            return 1;   //表示导出成功
        }
        public static int AllWeekDuxue()       //duxueFunction中获取导出路径
        {
            string rfilepath = "";


            OpenFileDialog dialog = new OpenFileDialog();
            //dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择需处理教务文件";
            dialog.Filter = "所有文件(*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                rfilepath = dialog.FileName;

            }
            else
                return 0;   //表示取消
            //string rfilepath = rpath;
            string wfilepath = "";          //写的路径

            if (readJiaowu(rfilepath) == -1)
            {
                MessageBox.Show("选错文件或所选教务信息文件已打开");
                return -1;
            }

            FolderBrowserDialog dialog1 = new FolderBrowserDialog();

            dialog1.Description = "请选择导出文件路径";

            if (dialog1.ShowDialog() == DialogResult.OK)
            {
                wfilepath = dialog1.SelectedPath + @"\";
            }
            else
                return 0;

            Information ii = head.next;
            int num = 2;
            int lieshu = 14;
            int realLieshu = 10;
            string[] headRow = { "1", "星期", "节", "课程名称", "地点", "推荐班级", "年级", "周", "任课教师", "人数", "教师情况", "学生情况（出勤率、迟到人数）", "教学效果", "意见和建议" };
            wfilepath += "/督学.xlsx";
            IWorkbook wb = null;
            try
            {
                FileStream fs = File.Create(wfilepath);
                wb = new HSSFWorkbook();
                ISheet sheet1 = wb.CreateSheet("sheet1");
                sheet1.CreateRow(0);
                for (int i = 0; i < lieshu; i++)       //创建表头
                {
                    sheet1.GetRow(0).CreateCell(i);
                    sheet1.GetRow(0).GetCell(i).SetCellValue(headRow[i]);

                }
                while (ii != null)
                {
                    sheet1.CreateRow(num - 1);
                    for (int i = 0; i < realLieshu; i++)
                    {

                        sheet1.GetRow(num - 1).CreateCell(i);
                    }
                    sheet1.GetRow(num - 1).GetCell(0).SetCellValue(num.ToString());
                    sheet1.GetRow(num - 1).GetCell(1).SetCellValue(switchDay(ii.day));
                    sheet1.GetRow(num - 1).GetCell(2).SetCellValue(ii.whichLesson);
                    sheet1.GetRow(num - 1).GetCell(3).SetCellValue(ii.lessenName);
                    sheet1.GetRow(num - 1).GetCell(4).SetCellValue(ii.whichPlace);
                    sheet1.GetRow(num - 1).GetCell(5).SetCellValue(ii.whichClass);
                    sheet1.GetRow(num - 1).GetCell(6).SetCellValue(ii.whichGrade);
                    sheet1.GetRow(num - 1).GetCell(7).SetCellValue(ii.whichWeek);
                    sheet1.GetRow(num - 1).GetCell(8).SetCellValue(ii.teacher);
                    sheet1.GetRow(num - 1).GetCell(9).SetCellValue(ii.personNum);
                    num++;
                    ii = ii.next;
                }

                Console.WriteLine(num);
                wb.Write(fs);

                fs.Close();
                wb.Close();
            }
            catch (Exception e)
            {
                return -1;      //表示失败
            }
            return 1;//表示正常运行
        }
        

    }

}
