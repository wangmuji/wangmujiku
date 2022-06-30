using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WF=WindowsFormsApp1;

namespace WindowsFormsApp1
{
    public partial class mainForm : Form
    {
        public mainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)      // 督学
        {
            
            duxueForm dx = new duxueForm();
            
            dx.Owner = this;
            this.Hide();
            dx.ShowDialog();
            
            
        }

        private void button2_Click(object sender, EventArgs e)      //工作量
        {
            gzlForm gzl = new gzlForm();

            gzl.Owner = this;
            this.Hide();
            gzl.ShowDialog();

        }

        private void button3_Click(object sender, EventArgs e)      //监考
        {
            jiankaoForm jk = new jiankaoForm();

            jk.Owner = this;
            this.Hide();
            jk.ShowDialog();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
    
}
