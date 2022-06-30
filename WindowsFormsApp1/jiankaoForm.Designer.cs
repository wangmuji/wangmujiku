
using System;
using System.Drawing;
using System.Windows.Forms;
namespace WindowsFormsApp1
{
    partial class jiankaoForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            this.Hide();
            this.Owner.Show();
            this.Dispose();
        }

        public class ComBox : ComboBox
        {
            public ComBox()
            {
                this.DropDownStyle = ComboBoxStyle.DropDownList;
                this.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
                this.DrawItem += new DrawItemEventHandler(ComBox_DrawItem);
            }

            void ComBox_DrawItem(object sender, DrawItemEventArgs e)
            {
                ComboBox cb = sender as ComboBox;
                if (e.Index < 0)
                {
                    return;
                }
                e.DrawBackground();
                e.DrawFocusRectangle();
                e.Graphics.DrawString(cb.GetItemText(cb.Items[e.Index]).ToString(), e.Font, new SolidBrush(e.ForeColor), e.Bounds.X, e.Bounds.Y + 3);
            }

        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.comBox1 = new WindowsFormsApp1.jiankaoForm.ComBox();
            this.comBox2 = new WindowsFormsApp1.jiankaoForm.ComBox();
            this.comBox3 = new WindowsFormsApp1.jiankaoForm.ComBox();
            this.comBox4 = new WindowsFormsApp1.jiankaoForm.ComBox();
            this.comBox5 = new WindowsFormsApp1.jiankaoForm.ComBox();
            this.comBox6 = new WindowsFormsApp1.jiankaoForm.ComBox();
            this.comBox7 = new WindowsFormsApp1.jiankaoForm.ComBox();
            this.comBox8 = new WindowsFormsApp1.jiankaoForm.ComBox();
            this.comBox9 = new WindowsFormsApp1.jiankaoForm.ComBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label11 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.label12 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.label13 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // comBox1
            // 
            this.comBox1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.comBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comBox1.FormattingEnabled = true;
            this.comBox1.Items.AddRange(new object[] {
            "2020",
            "2021",
            "2022",
            "2023",
            "2024",
            "2025",
            "2026",
            "2027",
            "2028",
            "2029",
            "2030",
            "2031",
            "2032",
            "2033",
            "2034",
            "2035"});
            this.comBox1.Location = new System.Drawing.Point(118, 124);
            this.comBox1.Name = "comBox1";
            this.comBox1.Size = new System.Drawing.Size(70, 26);
            this.comBox1.TabIndex = 0;
            this.comBox1.SelectedIndexChanged += new System.EventHandler(this.comBox1_SelectedIndexChanged);
            // 
            // comBox2
            // 
            this.comBox2.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.comBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comBox2.FormattingEnabled = true;
            this.comBox2.Location = new System.Drawing.Point(228, 124);
            this.comBox2.MaxDropDownItems = 7;
            this.comBox2.Name = "comBox2";
            this.comBox2.Size = new System.Drawing.Size(57, 26);
            this.comBox2.TabIndex = 0;
            this.comBox2.SelectedIndexChanged += new System.EventHandler(this.comBox2_SelectedIndexChanged);
            // 
            // comBox3
            // 
            this.comBox3.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.comBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comBox3.FormattingEnabled = true;
            this.comBox3.ItemHeight = 20;
            this.comBox3.Location = new System.Drawing.Point(323, 124);
            this.comBox3.MaxDropDownItems = 7;
            this.comBox3.Name = "comBox3";
            this.comBox3.Size = new System.Drawing.Size(57, 26);
            this.comBox3.TabIndex = 0;
            this.comBox3.SelectedIndexChanged += new System.EventHandler(this.comBox3_SelectedIndexChanged);
            // 
            // comBox4
            // 
            this.comBox4.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.comBox4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comBox4.FormattingEnabled = true;
            this.comBox4.ItemHeight = 20;
            this.comBox4.Items.AddRange(new object[] {
            "0",
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
            "12",
            "13",
            "14",
            "15",
            "16",
            "17",
            "18",
            "19",
            "20",
            "21",
            "22",
            "23"});
            this.comBox4.Location = new System.Drawing.Point(118, 173);
            this.comBox4.MaxDropDownItems = 7;
            this.comBox4.Name = "comBox4";
            this.comBox4.Size = new System.Drawing.Size(57, 26);
            this.comBox4.TabIndex = 0;
            this.comBox4.SelectedIndexChanged += new System.EventHandler(this.comBox4_SelectedIndexChanged);
            // 
            // comBox5
            // 
            this.comBox5.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.comBox5.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comBox5.FormattingEnabled = true;
            this.comBox5.ItemHeight = 20;
            this.comBox5.Location = new System.Drawing.Point(198, 173);
            this.comBox5.MaxDropDownItems = 7;
            this.comBox5.Name = "comBox5";
            this.comBox5.Size = new System.Drawing.Size(57, 26);
            this.comBox5.TabIndex = 0;
            this.comBox5.SelectedIndexChanged += new System.EventHandler(this.comBox5_SelectedIndexChanged);
            // 
            // comBox6
            // 
            this.comBox6.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.comBox6.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comBox6.FormattingEnabled = true;
            this.comBox6.ItemHeight = 20;
            this.comBox6.Location = new System.Drawing.Point(279, 173);
            this.comBox6.MaxDropDownItems = 7;
            this.comBox6.Name = "comBox6";
            this.comBox6.Size = new System.Drawing.Size(57, 26);
            this.comBox6.TabIndex = 0;
            this.comBox6.SelectedIndexChanged += new System.EventHandler(this.comBox6_SelectedIndexChanged);
            // 
            // comBox7
            // 
            this.comBox7.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.comBox7.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comBox7.FormattingEnabled = true;
            this.comBox7.ItemHeight = 20;
            this.comBox7.Location = new System.Drawing.Point(357, 173);
            this.comBox7.MaxDropDownItems = 7;
            this.comBox7.Name = "comBox7";
            this.comBox7.Size = new System.Drawing.Size(57, 26);
            this.comBox7.TabIndex = 0;
            this.comBox7.SelectedIndexChanged += new System.EventHandler(this.comBox7_SelectedIndexChanged);
            // 
            // comBox8
            // 
            this.comBox8.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.comBox8.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comBox8.Location = new System.Drawing.Point(118, 73);
            this.comBox8.Name = "comBox8";
            this.comBox8.Size = new System.Drawing.Size(242, 26);
            this.comBox8.TabIndex = 4;
            this.comBox8.SelectedIndexChanged += new System.EventHandler(this.comBox8_SelectedIndexChanged);
            // 
            // comBox9
            // 
            this.comBox9.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.comBox9.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comBox9.Location = new System.Drawing.Point(119, 314);
            this.comBox9.Name = "comBox9";
            this.comBox9.Size = new System.Drawing.Size(323, 26);
            this.comBox9.TabIndex = 4;
            this.comBox9.SelectedIndexChanged += new System.EventHandler(this.comBox8_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(194, 131);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(28, 19);
            this.label1.TabIndex = 1;
            this.label1.Text = "年";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(289, 131);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(28, 19);
            this.label2.TabIndex = 2;
            this.label2.Text = "月";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("宋体", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(386, 131);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(28, 19);
            this.label3.TabIndex = 3;
            this.label3.Text = "日";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(177, 176);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(16, 15);
            this.label4.TabIndex = 5;
            this.label4.Text = ":";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(261, 176);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(15, 15);
            this.label5.TabIndex = 6;
            this.label5.Text = "-";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.Location = new System.Drawing.Point(339, 176);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(16, 15);
            this.label6.TabIndex = 7;
            this.label6.Text = ":";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("宋体", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label7.Location = new System.Drawing.Point(115, 51);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(66, 19);
            this.label7.TabIndex = 8;
            this.label7.Text = "教研室";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("宋体", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label8.Location = new System.Drawing.Point(114, 102);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(47, 19);
            this.label8.TabIndex = 9;
            this.label8.Text = "日期";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("宋体", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label9.Location = new System.Drawing.Point(115, 153);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(47, 19);
            this.label9.TabIndex = 10;
            this.label9.Text = "时间";
            // 
            // label10
            // 
            this.label10.Font = new System.Drawing.Font("华康方圆体W7", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label10.Location = new System.Drawing.Point(35, 136);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(35, 129);
            this.label10.TabIndex = 11;
            this.label10.Text = "排除功能";
            this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("宋体", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button1.Location = new System.Drawing.Point(118, 217);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(80, 37);
            this.button1.TabIndex = 12;
            this.button1.Text = "添加";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("宋体", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label11.Location = new System.Drawing.Point(115, 286);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(142, 19);
            this.label11.TabIndex = 13;
            this.label11.Text = "已添加排除信息";
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("宋体", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button2.Location = new System.Drawing.Point(119, 362);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(80, 38);
            this.button2.TabIndex = 14;
            this.button2.Text = "删除";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("宋体", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button3.Location = new System.Drawing.Point(581, 235);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(102, 30);
            this.button3.TabIndex = 15;
            this.button3.Text = "生成监考表";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("华康方圆体W7", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label12.Location = new System.Drawing.Point(540, 163);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(193, 30);
            this.label12.TabIndex = 16;
            this.label12.Text = "监考分配功能";
            this.label12.Click += new System.EventHandler(this.label12_Click);
            // 
            // button4
            // 
            this.button4.Font = new System.Drawing.Font("宋体", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button4.Location = new System.Drawing.Point(264, 11);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(102, 31);
            this.button4.TabIndex = 17;
            this.button4.Text = "选择文件";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("宋体", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label13.Location = new System.Drawing.Point(115, 17);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(142, 19);
            this.label13.TabIndex = 18;
            this.label13.Text = "选择教研室文件";
            // 
            // jiankaoForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comBox1);
            this.Controls.Add(this.comBox2);
            this.Controls.Add(this.comBox3);
            this.Controls.Add(this.comBox4);
            this.Controls.Add(this.comBox5);
            this.Controls.Add(this.comBox6);
            this.Controls.Add(this.comBox7);
            this.Controls.Add(this.comBox8);
            this.Controls.Add(this.comBox9);
            this.Name = "jiankaoForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "生成监考表";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ComBox comBox1;
        private ComBox comBox2;
        private ComBox comBox3;
        private ComBox comBox4;
        private ComBox comBox5;
        private ComBox comBox6;
        private ComBox comBox7;
        private ComBox comBox8;
        private ComBox comBox9;
        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private Label label5;
        private Label label6;
        private Label label7;
        private Label label8;
        private Label label9;
        private Label label10;
        private Button button1;
        private Label label11;
        private Button button2;
        private Button button3;
        private Label label12;
        private Button button4;
        private Label label13;
    }
}