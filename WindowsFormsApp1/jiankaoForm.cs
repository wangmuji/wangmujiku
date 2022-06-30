using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace WindowsFormsApp1
{

    public partial class jiankaoForm : Form
    {
        static Office officeHead = new Office();        //头节点为哨兵，内无内容
        static Office officeRear = null;
        static Invigilation inviHead = new Invigilation();
        static Invigilation inviRear = null;
        static paichu pHead = new paichu();
        static paichu pRear = null;

        public jiankaoForm()
        {
            InitializeComponent();
            
        }
       
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void comBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comBox2.Items.Clear();
            comBox3.Items.Clear();
            comBox2.Text = "";
            comBox3.Text = "";
            this.comBox2.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "11",
            "12"});

        }

        private void comBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            comBox3.Text = "";
            comBox3.Items.Clear();
            int year = int.Parse(comBox1.Text);
            int month = int.Parse(comBox2.Text);
            if (month == 4 || month == 6 || month == 9 || month == 11)     //小月
            {
                for (int i = 1; i < 31; i++)
                {
                    this.comBox3.Items.Add(i.ToString());
                }
            }
            else if (month == 2)
            {
                if (year % 400 == 0 || (year % 4 == 0 && year % 100 != 0))       //闰年
                {
                    for (int i = 1; i < 30; i++)
                    {
                        this.comBox3.Items.Add(i.ToString());
                    }
                }
                else
                {
                    for (int i = 1; i < 29; i++)
                    {
                        this.comBox3.Items.Add(i.ToString());
                    }
                }
            }
            else        //大月
            {
                for (int i = 1; i < 32; i++)
                {
                    this.comBox3.Items.Add(i.ToString());
                }
            }

        }

        private void comBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

            comBox5.Text = "";
            comBox6.Text = "";
            comBox7.Text = "";
            comBox5.Items.Clear();
            comBox6.Items.Clear();
            comBox7.Items.Clear();
            for (int i = 0; i < 60; i++)
            {
                comBox5.Items.Add(i.ToString());
            }
        }


        private void setMinuteItem(int fh, int fm, int lh)
        {
            if (fh == lh)
            {
                for (int i = fm; i < 60; i++)
                {
                    comBox7.Items.Add(i.ToString());
                }
            }
            else
            {
                for (int i = 0; i < 60; i++)
                {
                    comBox7.Items.Add(i.ToString());
                }
            }
        }

        private void comBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            int firstH = int.Parse(comBox4.Text);
            int firstM = int.Parse(comBox5.Text);
            comBox6.Text = "";
            comBox7.Text = "";
            comBox6.Items.Clear();
            comBox7.Items.Clear();
            for (int i = firstH; i < 24; i++)
            {
                comBox6.Items.Add(i.ToString());
            }
        }

        private void comBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            int firstH = int.Parse(comBox4.Text);
            int firstM = int.Parse(comBox5.Text);
            int lastH = int.Parse(comBox6.Text);
            comBox7.Text = "";
            comBox7.Items.Clear();
            setMinuteItem(firstH, firstM, lastH);
        }

        private void comBox7_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comBox8_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        //
        //以下为监考功能函数
        //
        public static int readJYS(string filePath)
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

                row = sheet.GetRow(0);
                if (row != null)
                {
                    if (row.GetCell(0) == null)
                    {
                        return -1;
                    }
                    else
                    {
                        if (!row.GetCell(0).ToString().Equals("教研室"))          //判断文件依据
                        {
                            return -1;
                        }
                    }
                }
                Office office = null;
                for (int i = 1; i <= sheet.LastRowNum; i++)         //从第二行开始
                {
                    row = sheet.GetRow(i);  //读取当前行数据
                    if (row != null)
                    {

                        //LastCellNum 是当前行的总列数
                        for (int j = 0; j < row.LastCellNum; j++)        //
                        {
                            //读取该行的第j列数据
                            string value = string.Empty;
                            if (row.GetCell(j) != null)
                                value = row.GetCell(j).ToString();
                            else
                                continue;       //该行的这列为空，多半文件出错或是选错了文件
                            if (j == 0)
                            {
                                office = new Office();
                                office.setOfficeName(value);      //教研室名称
                                if (officeHead.next == null)
                                {
                                    officeHead.next = office;
                                    officeRear = office;
                                }
                                else        //在此处插入链表
                                {
                                    officeRear.next = office;
                                    officeRear = office;
                                }
                            }
                            else if (j == 1)
                            {
                                office.setPeopleNum(int.Parse(value));          //教研室人数
                            }
                            else if (j == 2)
                            {
                                office.setDirector(value);          //教研室主任
                            }


                        }
                        //Console.WriteLine(office.getOfficeName() + " " + office.getPeopleNum() + " " + office.getDirector());
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
        public static void JKs(Invigilation head)        //算监考人数
        {
            Invigilation inv = head.next;
            while (inv != null)
            {
                //   Console.WriteLine("aa");
                int d = inv.getStuNum();
                if (d < 70 && d > 0)
                    inv.setTeacNum(2);
                if (d >= 70 && d <= 100)
                    inv.setTeacNum(3);
                if (d > 100 && d <= 150)
                    inv.setTeacNum(4);
                if (d > 150 && d <= 180)
                    inv.setTeacNum(5);
                if (d > 180)
                    inv.setTeacNum(6);
                inv = inv.next;
            }
            inv = head.next;
            while (inv != null)
            {
                Console.WriteLine(inv.getDate() + " " + inv.getTeacNum());
                inv = inv.next;
            }

        }
        public static int readJK(string filePath)
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

                int offset = 0;
                int NUM = 1;
                row = sheet.GetRow(0);
                if (row != null)
                {
                    if (row.GetCell(0) == null)
                    {
                        return -1;
                    }
                    else
                    {
                        if (!row.GetCell(0).ToString().Contains("监考"))          //判断文件依据
                        {
                            return -1;
                        }
                    }
                }

                Invigilation invi = null;
                //LastRowNum 是当前表的总行数-1（注意）

                for (int i = 2; i <= sheet.LastRowNum; i++)         //从第二行开始
                {
                    row = sheet.GetRow(i);  //读取当前行数据
                    if (row != null)
                    {
                        //LastCellNum 是当前行的总列数
                        for (int j = 0; j < row.LastCellNum; j++)        //
                        {
                            //读取该行的第j列数据
                            string value = string.Empty;
                            if (row.GetCell(j) != null)
                            {

                                value = row.GetCell(j).ToString();


                            }
                            else
                                continue;       //该行的这列为空，多半文件出错或是选错了文件
                            if (j == 0)
                            {
                                //if(!value.Contains("月"))
                                //value = DateTime.FromOADate(double.Parse(value)).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo);
                                invi = new Invigilation();
                                invi.setDate(row.GetCell(j).DateCellValue.ToString().Split(' ')[0]);
                                if (inviHead.next == null)
                                {
                                    inviHead.next = invi;
                                    inviRear = invi;
                                }
                                else        //在此处插入链表
                                {
                                    inviRear.next = invi;
                                    inviRear = invi;
                                }
                            }
                            else if (j == 1)
                            {
                                invi.setTime(value);
                            }
                            else if (j == 6)
                            {
                                //
                                if (row.GetCell(j) != null && value != "")
                                {
                                    invi.setStuNum(int.Parse(value));

                                }


                            }
                            else if (j == 7)
                            {
                                invi.setPlace(value);
                            }
                        }

                       
                    }
                }
            }
            catch (Exception e)
            {
                //只在Debug模式下才输出
                Console.WriteLine("error");
                return -1;
            }

            while (inviHead.next != null)
            {
                Console.WriteLine(inviHead.next.getDate() + " " + inviHead.next.getTime() + " " + inviHead.getStuNum() + " " + inviHead.next.getPlace());
                inviHead = inviHead.next;
            }

            return 1;
        }
        public static int writeJK(string filePath)
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

                int offset = 0;
                int NUM = 1;
                row = sheet.GetRow(0);
                if (row != null)
                {
                    if (row.GetCell(0) == null)
                    {
                        return -1;
                    }
                    else
                    {
                        if (!row.GetCell(0).ToString().Contains("监考"))          //判断文件依据
                        {
                            return -1;
                        }
                    }
                }

                Invigilation invi = null;
                //LastRowNum 是当前表的总行数-1（注意）

                for (int i = 2; i <= sheet.LastRowNum; i++)         //从第二行开始
                {
                    row = sheet.GetRow(i);  //读取当前行数据
                    if (row != null)
                    {
                        //LastCellNum 是当前行的总列数
                        for (int j = 0; j < row.LastCellNum; j++)        //
                        {
                            //读取该行的第j列数据
                            string value = string.Empty;
                            if (row.GetCell(j) != null)
                            {

                                value = row.GetCell(j).ToString();


                            }
                            else
                                continue;       //该行的这列为空，多半文件出错或是选错了文件
                            if (j == 0)
                            {
                                //if(!value.Contains("月"))
                                //value = DateTime.FromOADate(double.Parse(value)).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo);
                                invi = new Invigilation();
                                invi.setDate(row.GetCell(j).DateCellValue.ToString().Split(' ')[0]);
                                if (inviHead.next == null)
                                {
                                    inviHead.next = invi;
                                    inviRear = invi;
                                }
                                else        //在此处插入链表
                                {
                                    inviRear.next = invi;
                                    inviRear = invi;
                                }
                            }
                            else if (j == 1)
                            {
                                invi.setTime(value);
                            }
                            else if (j == 6)
                            {
                                //
                                if (row.GetCell(j) != null && value != "")
                                {
                                    invi.setStuNum(int.Parse(value));

                                }


                            }
                            else if (j == 7)
                            {
                                invi.setPlace(value);
                            }
                        }

                        //Console.WriteLine("\n");
                    }
                }
                JKs(inviHead);
                JKE(inviHead, officeHead.next);
                Invigilation ia = inviHead.next;
                for (int i = 2; i <= sheet.LastRowNum; i++)
                {
                    row = sheet.GetRow(i);
                    string s = " ";
                    if (ia.getDirector() != null)
                        s = ia.getDirector().ToString();
                    row.GetCell(9).SetCellValue(s);
                    //  Console.WriteLine(ia.getDirector());
                    ia = ia.next;

                    if (ia == null)
                        break;
                }
                FileStream file = File.Create(filePath);
                wk.Write(file);
                file.Close();
                wk.Close();
                MessageBox.Show("分配成功");
            }
            catch (Exception e)
            {
                //只在Debug模式下才输出
                Console.WriteLine("error");
                MessageBox.Show("分配失败");
                return -1;
            }

         

            return 1;
        }
        public void setCom8()          //给教研室名称下拉菜单赋值
        {
            Office oo = officeHead;
            while (oo.next != null)
            {
                this.comBox8.Items.Add(oo.next.getOfficeName());
                oo = oo.next;
            }
        }
        public static int addPaichu(string date, string time, string officeName)
        {
            paichu pp = new paichu();
            pp.setDate(date);
            pp.setTime(time);
            pp.setOfficeName(officeName);
            if (pHead.next == null)
            {
                pHead.next = pp;
                pRear = pHead.next;

            }
            else
            {
                pRear.next = pp;
                pRear = pp;
            }
            return 1;
        }

        public static int getTimeNumber(string time)
        {
            int ntime = 0;
            ntime = int.Parse(time.Split(':')[0]) * 60 + int.Parse(time.Split(':')[1]);
            return ntime;
        }

        public static int judgePaichu(string date, string time, string officeName)       //判断重复时间段
        {
            int flag = 1;
            int oldfirst = 0;
            int oldlast = 0;
            int newfirst = getTimeNumber(time.Split('-')[0]);
            int newlast = getTimeNumber(time.Split('-')[1]);
            paichu pp = pHead;
            while (pp.next != null)
            {
                if (date.Equals(pp.next.getDate()) && officeName.Equals(pp.next.getOfficeName()))      //同一天存在同一教研室
                {
                    oldfirst = getTimeNumber(pp.next.getTime().Split('-')[0]);
                    oldlast = getTimeNumber(pp.next.getTime().Split('-')[1]);
                    if (oldfirst > newlast || oldlast < newfirst)  //作为新行添加
                    {
                    }
                    else
                    {
                        flag = -1;          //存在重复时间段
                    }
                }
                pp = pp.next;
            }


            return flag;
        }
        public void updateCom9()
        {
            this.comBox9.Text = "";
            this.comBox9.Items.Clear();
            paichu pp = pHead;
            while (pp.next != null)
            {
                comBox9.Items.Add(pp.next.getOfficeName() + "," + pp.next.getDate() + "," + pp.next.getTime());
                pp = pp.next;
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            if (comBox8.Text != "" && comBox7.Text != "" && comBox6.Text != "" && comBox5.Text != "" && comBox4.Text != "" && comBox3.Text != "" && comBox2.Text != "" && comBox1.Text != "")
            {

                string date = comBox1.Text + "/" + comBox2.Text + "/" + comBox3.Text;
                string time = comBox4.Text + ":" + comBox5.Text + "-" + comBox6.Text + ":" + comBox7.Text;
                string officeName = comBox8.Text;
                if (judgePaichu(date, time, officeName) == 1)
                {
                    addPaichu(date, time, officeName);
                    comBox1.Text = "";
                    comBox2.Text = "";
                    comBox3.Text = "";
                    comBox4.Text = "";
                    comBox5.Text = "";
                    comBox6.Text = "";
                    comBox7.Text = "";
                    comBox8.Text = "";
                    comBox1.Items.Clear();
                    comBox2.Items.Clear();
                    comBox3.Items.Clear();
                    comBox4.Items.Clear();
                    comBox5.Items.Clear();
                    comBox6.Items.Clear();
                    comBox7.Items.Clear();
                    comBox8.Items.Clear();
                    for (int i = 2020; i < 2036; i++)
                    {
                        comBox1.Items.Add(i.ToString());
                    }
                    for (int i = 0; i < 24; i++)
                    {
                        comBox4.Items.Add(i.ToString());
                    }
                    setCom8();
             //       addPaichu(officeName, date, time);
                    MessageBox.Show("成功添加:" + officeName + " " + date + " " + time);
                    updateCom9();
                    //刷新combo9-》已添加项
                }
                else
                {
                    MessageBox.Show("该教研室重复添加时间段，添加失败");
                }
            }
            else
            {
                MessageBox.Show("条件未填写完毕，添加失败");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string officeName;
            string date;
            string time;
            if (comBox9.Text != "")
            {
                officeName = comBox9.Text.Split(',')[0];
                date = comBox9.Text.Split(',')[1];
                time = comBox9.Text.Split(',')[2];
                paichu pp = pHead;
                while (pp.next != null)
                {
                    if (pp.next.getOfficeName().Equals(officeName) && pp.next.getDate().Equals(date) && pp.next.getTime().Equals(time))
                    {
                        MessageBox.Show("成功删除" + pp.next.getOfficeName() + "," + pp.next.getDate() + "," + pp.next.getTime());
                        pp.next = pp.next.next;     //将pp.next从链表中删除
                        updateCom9();

                        break;
                    }
                    pp = pp.next;
                }
            }

        }
        public static bool expect(string date, string time, Office office)//查此时间点是否被排除
        {
            paichu head = pHead.next;
            while (head != null)
            {
                String[] array = head.getTime().Split('-');
                String[] sarray = time.Split('-');
                int []a = new int[2];
                int[] sa = new int[2];
                a[0] = getTimeNumber(array[0]);
                a[1] = getTimeNumber(array[1]);
                sa[0] = getTimeNumber(sarray[0]);
                sa[1] = getTimeNumber(sarray[1]);
                if (office.getOfficeName() == head.getOfficeName())
                {
                    if (date == head.getDate())
                    {
                        if (sa[0] >= a[0] && sa[0] <= a[1] || sa[1] >= a[0] && sa[1] <= a[1])
                        {
                            return false;
                        }
                    }
                }
                head = head.next;
            }
            return true;//没被排除
        }
        public static void JKE(Invigilation ihead, Office ohead)                        //
        {
            Invigilation inv = ihead.next;
            Office Of = ohead;
            while (inv != null)
            {
                if (Of == null)
                    Of = ohead;
                while (Of != null && inv.getTeacNum() > inv.getTeacHadNum())
                {
                    //   Console.WriteLine("3333");
                    if (expect(inv.getDate(), inv.getTime(), Of))
                    {
                        int x1 = 1;
                        switch (inv.getTeacNum())
                        {
                            case 2: x1 = 2; break;
                            case 3: x1 = 3; break;
                            case 4: x1 = 2; break;
                            case 5: x1 = 3; break;
                            case 6: x1 = 3; break;
                            case 7: x1 = 4; break;
                        }
                        int x = inv.getTeacNum() - inv.getTeacHadNum();//此门考试还需要的人数                      
                                                                       //   Console.WriteLine(inv.getTeacNum());
                        if (x >= x1)
                        {
                            if (inv.getDirector() == null)
                            {
                                inv.setDirector(inv.getDirector() + Of.getDirector() + (x1).ToString() + "人");
                            }
                            else
                            inv.setDirector(inv.getDirector() + "+" + Of.getDirector() + (x1).ToString() + "人");
                            inv.setTeacHadNum(inv.getTeacHadNum() + x1);

                        }
                        else
                        {
                            inv.setDirector(inv.getDirector() + "+" + Of.getDirector() + (x1).ToString() + "人");
                            inv.setTeacHadNum(inv.getTeacHadNum() + x);

                            break;//这里说明这门考试的人数分配完成

                        }
                        //       Console.WriteLine(inv.getTeacNum()+"      "+inv.getDirector() );

                        //    Of.setPassNum(Of.getPassNum() + 1);
                    }
                    Of = Of.next;
                }
                inv = inv.next;//Console.WriteLine(  "444444");
            }
        }
        public static void gouzao(string date, string time, Office office)//构造排除信息
        {
            paichu p = new paichu();
            p.setDate(date);
            p.setTime(time);
            p.setOfficeName(office.getOfficeName());
            if (pHead.next == null)
            {
                pHead.next = p;
                pRear = p;
            }
            else
            {
                pRear.next = null;
                pRear = p;
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            string rfilepath = "";
            OpenFileDialog dialog = new OpenFileDialog();
            //dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择需处理监考文件";
            dialog.Filter = "所有文件(*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                rfilepath = dialog.FileName;
            }
            if(writeJK(rfilepath)==-1)
            {
                MessageBox.Show("选择文件错误或文件打开中");
            }
        }



        public class Office
        {
            string officeName;      //教研室名字
            string director;        //教研室主任
            int peopleNum;          //教研室人数
            int taskNum;            //任务应分配人数
            int passNum;            //已完成任务数

            public Office next;
            public Office()
            {
                officeName = "教研室";
                director = "999";
                peopleNum = 0;
                taskNum = 0;
                passNum = 0;
                next = null;
            }
            public string getOfficeName() { return this.officeName; }
            public string getDirector() { return this.director; }
            public int getPeopleNum() { return this.peopleNum; }
            public int getPassNum() { return this.passNum; }
            public int getTaskNum() { return this.taskNum; }
            public void setTaskNum(int task) { this.taskNum = task; }
            public void setPassNum(int pass) { this.passNum = pass; }
            public void setDirector(string di) { this.director = di; }
            public void setPeopleNum(int people) { this.peopleNum = people; }
            public void setOfficeName(string name) { this.officeName = name; }
        }

        public class Invigilation       //监考类
        {
            string time;        //考试时间
            string date;            //考试日期
            string place;       //考试地点
            int stuNum;     //考生人数
            int teacNum;     //需要的监考老师人数
            int teacHadNum;     //已经有的监考老师人数
            string director;//主任名称和该科室的分配信息(可为复数)
            public Invigilation next;
            public Invigilation()
            {
                date = "2000.1.1";
                time = "0:0-23:59";
                place = "place";
                stuNum = 0;
                teacNum = 0;
                teacHadNum = 0;
                director = null;
                next = null;
            }
            public string getTime() { return this.time; }
            public string getPlace() { return this.place; }
            public int getStuNum() { return this.stuNum; }
            public string getDate() { return this.date; }
            public int getTeacNum() { return this.teacNum; }

            public int getTeacHadNum() { return this.teacHadNum; }
            public string getDirector() { return this.director; }
            public void setDate(string date) { this.date = date; }
            public void setTime(string time) { this.time = time; }
            public void setStuNum(int stu) { this.stuNum = stu; }
            public void setPlace(string place) { this.place = place; }
            public void setTeacNum(int teac) { this.teacNum = teac; }
            public void setTeacHadNum(int teacHad) { this.teacHadNum = teacHad; }
            public void setDirector(string director) { this.director = director; }
            //包含监考信息
        }

        public class paichu
        {
            string date;        //排除的日期    ,格式：yyyy/mm/dd      ps:mm不以0开头，为1、2...10，不为01、02...dd同理
            string time;        //排除的时间段   ,格式：xx:xx-xx:xx
            string officeName;  //排除的教研室名字
            public paichu next;
            public paichu()         //构造函数
            {
                date = "2000.1.1";
                time = "0:0-23:59";
                officeName = "office";
                next = null;
            }
            public string getOfficeName() { return this.officeName; }
            public string getDate() { return this.date; }
            public string getTime() { return this.time; }
            public void setDate(string date) { this.date = date; }
            public void setTime(string time) { this.time = time; }
            public void setOfficeName(string name) { this.officeName = name; }
            //排除信息，功能：记录排除教研室信息
        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            string rfilepath = "";

            OpenFileDialog dialog = new OpenFileDialog();
            //dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择教研室文件";
            dialog.Filter = "所有文件(*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                rfilepath = dialog.FileName;

            }
            readJYS(rfilepath);
            setCom8();
        }
    }
}
