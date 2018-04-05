using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;

namespace 成本效益评价技术第二版
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            //dataGridView1.RowsAdded += new DataGridViewRowsAddedEventHandler(dgvEvent);
            //dataGridView1.RowsRemoved -= new DataGridViewRowsRemovedEventHandler(dgvEvent);
        }

        //private void dgvEvent(object sender, DataGridViewRowsAddedEventArgs e)
        //{
        //    for (int x = 0; x < dataGridView1.Rows.Count; x++)
        //    {
        //        dataGridView1.AllowUserToAddRows = true;
        //        if (x > 7)
        //        {
        //            dataGridView1.AllowUserToAddRows = false;
        //            MessageBox.Show("收入情况已输入完毕！");
        //        }
        //        //else
        //        //{
        //        //    dataGridView1.AllowUserToAddRows = true;
        //        //}
        //    }
        //    //dataGridView1.AllowUserToAddRows = false;
        //    //if (dataGridView1.RowCount < 5)
        //    //{
        //    //    dataGridView1.AllowUserToAddRows = true;
        //    //    //MessageBox.Show("收入情况已输入完毕！");
        //    //}                                                
        //    //else
        //    //{
        //    //    //MessageBox.Show("收入情况已输入完毕！");
        //    //    dataGridView1.AllowUserToAddRows = true;
                
        //    //}
        //}

        //计算之前首先判断是否为空
        private void button1_Click(object sender, EventArgs e)
        {
            int year,beginYear;
            double touRu;
            String tieXianLv="";


            //项目运行年数，项目初期投入（需判断是否是数值形式），以及项目起始年份，贴现率不能为空
            if (int.TryParse(UDYunXing.Text, out year) && int.TryParse(UDBeginYear.Text, out beginYear) && double.TryParse(tbTouRu.Text, out touRu))
            {
                if (tbTieXianLv.Text == "")
                {
                    MessageBox.Show("请输入贴现率！");
                }
                else
                {
                    for (int a = 1; a <= year; a++)
                    {
                        if (dataGridView1.RowCount==year)
                        {
                            //开始计算各项属性值
                            tieXianLv = tbTieXianLv.Text;
                            calculate();
                        }
                        else
                        {
                            MessageBox.Show("请将所有年份收入情况输入完毕！");
                            break;
                        }
                    }
                       
                }
            }
            else
            {
                MessageBox.Show("请正确输入各项属性值！");
            }


        }


        //计算各项属性函数
        public void calculate()
        {
            //项目运行年数
            int year;
            year =int.Parse( UDYunXing.Text);

            //定义每年收入的数组
            ArrayList shouRu = new ArrayList();
            //初期投入
            double touRu;
            touRu = double.Parse(tbTouRu.Text);
            shouRu.Add(-touRu);
           
            //将每年收入情况输入
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                shouRu.Add(dataGridView1.Rows[i].Cells["Column2"].Value);
            }

            //贴现率
            String s;
            double lv=0;
            s = tbTieXianLv.Text;
            int pos = s.IndexOf("%");
            int length = s.Length;
            if (pos == -1)//如果没有百分号
            {
                lv = double.Parse(s);//将字符串转换为double类型
            }
            else if (pos < length)//如果存在百分号
            {
                lv = (double.Parse(s.Substring(0,pos)))/100;
            }

            //成本效益评价:

            //净利润
            jingLiRun.Text = ""+NP(year,shouRu);
            //回收期
            huiShouQi.Text = "" + PP(year,shouRu);
            //投资回报率
            huiBaoLv.Text = ROI(year, shouRu);
            //净现值
            jingXianZhi.Text = NPV(year, shouRu,lv);
        }


        //净利润
        public double NP(int year, ArrayList p)
        {
            double np = 0;
            for (int a = 0; a <=year; a++)
            {
                np += (double.Parse(p[a].ToString()));
            }
            return np;
        }

        //回收期
        public int PP(int year, ArrayList p)
        {
            double pp = 0;
            int a;
            for (a = 0; a <= year; a++)
            {
                pp += (double.Parse(p[a].ToString()));
                if (pp >= 0)
                    break;
            }
            return a;
        }

        //投资回报率
        public String ROI(int year, ArrayList p)
        {
            //净利润
            double np = NP(year, p);

            double roi = 0;
            roi = np / (-(double.Parse(p[0].ToString()))*year);
            return roi.ToString("P");
        }

        //净现值
        public String NPV(int year, ArrayList p, double lv)
        {
            //贴现因子（存储在数组中）
            ArrayList f = new ArrayList();
            f.Add(1);
            for (int a = 1; a <= year; a++)
            {
                f.Add(1/(Math.Pow((1+lv),a)));
            }

            double z = 0;
            String s="";
            for (int a = 0; a <= year; a++)
            {
                s += "第" + a + "年的净现值为：" + ((double.Parse(p[a].ToString())) * (double.Parse(f[a].ToString()))).ToString("f4") + "元\r\n\r\n";
                z += (double.Parse(p[a].ToString())) * (double.Parse(f[a].ToString()));
            }
            s += "总净现值为：" + z.ToString("f4") + "元";
            return s;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.UserAddedRow += new DataGridViewRowEventHandler(dataGridView1_UserAddedRow);
        }
        private int m_AddCount = 0;
        private void dataGridView1_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            m_AddCount++;
            this.dataGridView1.Rows[m_AddCount-1].Cells["Column1"].Value = UDBeginYear.Value + m_AddCount - 1;
            if (m_AddCount == UDYunXing.Value) dataGridView1.AllowUserToAddRows = false;
        }


    }
}

