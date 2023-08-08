using System;
using System.Data;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AutoUpDataBoss
{
    public partial class AutoUpDataSMS : Form
    {
        public AutoUpDataSMS()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            timer1.Enabled = true;
            label3.Text = "已开启";
            label3.ForeColor = System.Drawing.Color.Blue;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            label3.Text = "已关闭";
            label3.ForeColor = System.Drawing.Color.Red;
        }

        public void output(string good)
        {
            listBox1.Items.Add(good);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                listBox4.Items.Add("人工日报开始：" + DateTime.Now.ToString());
                DataProcessing.baobiao1();
                DataProcessing.baobiao2();
             //   DataProcessing.baobiao4();
                DataProcessing.baobiao5();
             //   DataProcessing.baobiao6();
              //  DataProcessing.baobiao7();
              //  DataProcessing.baobiao8();
                DataProcessing.baobiao9();
                DataProcessing.baobiao10();
                DataProcessing.baobiao11();
                DataProcessing.baobiao12();
                DataProcessing.baobiao13();
                listBox4.Items.Add("人工日报结束：" + DateTime.Now.ToString());
            }
            catch (Exception ex)
            {
                Error err = new Error();
                err.LogError("错误信息", ex);
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                listBox4.Items.Add("人工周报开始：" + DateTime.Now.ToString());
                DataProcessing.zbaobiao1();
                DataProcessing.zbaobiao2();
                DataProcessing.zbaobiao3();
                DataProcessing.zbaobiao4();
                DataProcessing.zbaobiao5();
                DataProcessing.zbaobiao6();
                DataProcessing.zbaobiao7();
                DataProcessing.zbaobiao8();
                DataProcessing.zbaobiao9();
                DataProcessing.zbaobiao10();
                DataProcessing.zbaobiao12();
                DataProcessing.zbaobiao13();
                DataProcessing.zbaobiao14();
                DataProcessing.zbaobiao15();
                listBox4.Items.Add("人工周报结束：" + DateTime.Now.ToString());
            }
            catch (Exception ex)
            {
                Error err = new Error();
                err.LogError("错误信息", ex);
            }
        }
        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                listBox4.Items.Add("月报数量开始：" + DateTime.Now.ToString());
                DataProcessing.ybaobiao1();
                DataProcessing.ybaobiao2();
                DataProcessing.ybaobiao3();
                DataProcessing.ybaobiao4();
                DataProcessing.ybaobiao5();
                DataProcessing.ybaobiao7();
                DataProcessing.ybaobiao8();
                DataProcessing.ybaobiao9();
                DataProcessing.ybaobiao10();
                DataProcessing.ybaobiao11();
                DataProcessing.ybaobiao12();
                DataProcessing.ybaobiao13();
                DataProcessing.ybaobiao14();
                DataProcessing.ybaobiao15();
                DataProcessing.ybaobiao16();
                DataProcessing.ybaobiao17();
                DataProcessing.ybaobiao18();
                listBox4.Items.Add("月报数量结束：" + DateTime.Now.ToString());
                listBox4.Items.Add("经分数量开始：" + DateTime.Now.ToString());
                DataProcessing.yjbaobiao2();
                DataProcessing.yjbaobiao3();
                DataProcessing.yjbaobiao5();
                DataProcessing.yjbaobiao6();
                DataProcessing.yjbaobiao7();
                DataProcessing.yjbaobiao8();
                DataProcessing.yjbaobiao10();
                DataProcessing.yjbaobiao13();
                DataProcessing.yjbaobiao14();
                DataProcessing.yjbaobiao15();
                DataProcessing.yjbaobiao16();
                DataProcessing.yjbaobiao17();
                DataProcessing.yjbaobiao19();
                listBox4.Items.Add("经分数量结束：" + DateTime.Now.ToString());
            }
            catch (Exception ex)
            {
                Error err = new Error();
                err.LogError("错误信息", ex);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                listBox4.Items.Add("月报金额开始：" + DateTime.Now.ToString());
                DataProcessing.ybaobiao19();
                DataProcessing.ybaobiao20();
                DataProcessing.ybaobiao21();
                DataProcessing.ybaobiao22();
                DataProcessing.ybaobiao23();
                listBox4.Items.Add("月报金额结束：" + DateTime.Now.ToString());
                listBox4.Items.Add("经分金额开始：" + DateTime.Now.ToString());
                DataProcessing.yjbaobiao1();
                DataProcessing.yjbaobiao4();
                DataProcessing.yjbaobiao9();
                DataProcessing.yjbaobiao11();
                DataProcessing.yjbaobiao12();
                DataProcessing.yjbaobiao18();
                listBox4.Items.Add("经分金额结束：" + DateTime.Now.ToString());
            }
            catch (Exception ex)
            {
                Error err = new Error();
                err.LogError("错误信息", ex);
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            string dtt = DateTime.Now.ToString("HHmmss");// 得到 hour minute second  如果等于某个值就开始执行某个程序。
            if (dtt == "070000")//每天7:00:00 开始执行  071500  7:15:00
            {
                try
                {
                    if (DateTime.Now.ToString("dd") == "21") //每月21号执行财务月报
                    {
                        DataProcessing.ybaobiao24();
                        DataProcessing.ybaobiao25();
                    }
                }
                catch (Exception ex)
                {
                    label10.Text = "0";
                    label13.Text = (Int32.Parse(label13.Text) + 1).ToString();
                    Error err = new Error();
                    err.LogError("错误信息", ex);
                }
                //try
                //{
                //    if (DateTime.Now.ToString("dd") == "01") //每月一号执行复通数量类
                //    {
                //        DataProcessing.cbaobiao1();
                //        label23.Text = "已触发";
                //        label21.Text = DataProcessing.cbaobiao2();
                //    }
                //    else
                //    {
                //        label23.Text = "未触发";
                //    }
                //}
                //catch (Exception ex)
                //{
                //    label23.Text = "出错";
                //    Error err = new Error();
                //    err.LogError("错误信息", ex);
                //}
                try
                {
                    if (DateTime.Now.ToString("dd") == "15") //每月15号执行各业务新增欠费金额统计
                    {
                        listBox3.Items.Add("新增欠费金额开始：" + DateTime.Now.ToString());
                        DataProcessing.ybaobiao26();
                        listBox3.Items.Add("新增欠费金额结束：" + DateTime.Now.ToString());
                        label8.Text = (Int32.Parse(label8.Text) + 1).ToString();
                        label19.Text = (Int32.Parse(label19.Text) + 1).ToString();
                    }
                }
                catch (Exception ex)
                {
                    label10.Text = "0";
                    label13.Text = (Int32.Parse(label13.Text) + 1).ToString();
                    Error err = new Error();
                    err.LogError("错误信息", ex);
                }
                try
                {
                    listBox1.Items.Add("日报开始：" + DateTime.Now.ToString());
                    DataProcessing.baobiao1();
                    DataProcessing.baobiao2();
                 //   DataProcessing.baobiao4();
                    DataProcessing.baobiao5();
                 //   DataProcessing.baobiao6();
                  //  DataProcessing.baobiao7();
                  //  DataProcessing.baobiao8();
                    DataProcessing.baobiao9();
                    DataProcessing.baobiao10();
                    DataProcessing.baobiao11();
                    DataProcessing.baobiao12();
                    DataProcessing.baobiao13();
                    listBox1.Items.Add("日报结束：" + DateTime.Now.ToString());
                    label8.Text = (Int32.Parse(label8.Text) + 1).ToString();
                    label17.Text = (Int32.Parse(label18.Text) + 1).ToString();
                    label9.Text = "1";
                }
                catch (Exception ex)
                {
                    label9.Text = "0";
                    label13.Text = (Int32.Parse(label13.Text) + 1).ToString();
                    Error err = new Error();
                    err.LogError("错误信息", ex);
                }
                try
                {
                    if (DateTime.Now.DayOfWeek.ToString() == "Wednesday")
                    {
                        listBox2.Items.Add("欠停周报开始：" + DateTime.Now.ToString());
                        DataProcessing.zbaobiao16();
                        listBox2.Items.Add("欠停周报结束：" + DateTime.Now.ToString());
                        label8.Text = (Int32.Parse(label8.Text) + 1).ToString();
                        label18.Text = (Int32.Parse(label18.Text) + 1).ToString();
                        label10.Text = "1";
                    }
                }
                catch (Exception ex)
                {
                    label10.Text = "0";
                    label13.Text = (Int32.Parse(label13.Text) + 1).ToString();
                    Error err = new Error();
                    err.LogError("错误信息", ex);
                }
                try
                {
                    if (DateTime.Now.DayOfWeek.ToString() == "Friday")
                    {
                        listBox2.Items.Add("周报开始：" + DateTime.Now.ToString());
                        DataProcessing.zbaobiao1();
                        DataProcessing.zbaobiao2();
                        DataProcessing.zbaobiao3();
                        DataProcessing.zbaobiao4();
                        DataProcessing.zbaobiao5();
                        DataProcessing.zbaobiao6();
                        DataProcessing.zbaobiao7();
                        DataProcessing.zbaobiao8();
                        DataProcessing.zbaobiao9();
                        DataProcessing.zbaobiao10();
                        DataProcessing.zbaobiao12();
                        DataProcessing.zbaobiao13();
                        DataProcessing.zbaobiao14();
                        DataProcessing.zbaobiao15();
                        listBox2.Items.Add("周报结束：" + DateTime.Now.ToString());
                        label8.Text = (Int32.Parse(label8.Text) + 1).ToString();
                        label18.Text = (Int32.Parse(label18.Text) + 1).ToString();
                        label10.Text = "1";
                    }
                }
                catch (Exception ex)
                {
                    label10.Text = "0";
                    label13.Text = (Int32.Parse(label13.Text) + 1).ToString();
                    Error err = new Error();
                    err.LogError("错误信息", ex);
                }
                try
                {
                    if (DateTime.Now.ToString("dd") == "01") //每月一号执行数量类
                    {
                        listBox3.Items.Add("月报数量开始：" + DateTime.Now.ToString());
                        DataProcessing.ybaobiao1();
                        DataProcessing.ybaobiao2();
                        DataProcessing.ybaobiao3();
                        DataProcessing.ybaobiao4();
                        DataProcessing.ybaobiao5();
                        DataProcessing.ybaobiao7();
                        DataProcessing.ybaobiao8();
                        DataProcessing.ybaobiao9();
                        DataProcessing.ybaobiao10();
                        DataProcessing.ybaobiao11();
                        DataProcessing.ybaobiao12();
                        DataProcessing.ybaobiao13();
                        DataProcessing.ybaobiao14();
                        DataProcessing.ybaobiao15();
                        DataProcessing.ybaobiao16();
                        DataProcessing.ybaobiao17();
                        DataProcessing.ybaobiao18();
                        listBox3.Items.Add("月报数量结束：" + DateTime.Now.ToString());
                        listBox3.Items.Add("经分数量开始：" + DateTime.Now.ToString());
                        DataProcessing.yjbaobiao2();
                        DataProcessing.yjbaobiao3();
                        DataProcessing.yjbaobiao5();
                        DataProcessing.yjbaobiao6();
                        DataProcessing.yjbaobiao7();
                        DataProcessing.yjbaobiao8();
                        DataProcessing.yjbaobiao10();
                        DataProcessing.yjbaobiao13();
                        DataProcessing.yjbaobiao14();
                        DataProcessing.yjbaobiao15();
                        DataProcessing.yjbaobiao16();
                        DataProcessing.yjbaobiao17();
                        DataProcessing.yjbaobiao19();
                        listBox3.Items.Add("经分数量结束：" + DateTime.Now.ToString());
                        label8.Text = (Int32.Parse(label8.Text) + 1).ToString();
                        label19.Text = (Int32.Parse(label19.Text) + 1).ToString();
                        label11.Text = "1";

                    }
                }
                catch (Exception ex)
                {
                    label11.Text = "0";
                    label13.Text = (Int32.Parse(label13.Text) + 1).ToString();
                    Error err = new Error();
                    err.LogError("错误信息", ex);
                }
                //  if (DateTime.Now.ToString("dd") == "03" || DateTime.Now.ToString("dd") == "04") //每月三号或者四号号执行金额类
                //  {
                //listBox3.Items.Add("月报金额开始：" + DateTime.Now.ToString());
                //DataProcessing.ybaobiao19();
                //DataProcessing.ybaobiao20();
                //DataProcessing.ybaobiao21();
                //DataProcessing.ybaobiao22();
                //DataProcessing.ybaobiao23();
                //listBox3.Items.Add("月报金额结束：" + DateTime.Now.ToString());
                //listBox3.Items.Add("经分金额开始：" + DateTime.Now.ToString());
                //DataProcessing.yjbaobiao1();
                //DataProcessing.yjbaobiao4();
                //DataProcessing.yjbaobiao9();
                //DataProcessing.yjbaobiao11();
                //DataProcessing.yjbaobiao12();
                //DataProcessing.yjbaobiao18();
                //listBox3.Items.Add("经分金额结束：" + DateTime.Now.ToString());
                //   }

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                listBox4.Items.Add("财务月报开始：" + DateTime.Now.ToString());
                DataProcessing.ybaobiao24();
                DataProcessing.ybaobiao25();
                listBox4.Items.Add("财务月报结束：" + DateTime.Now.ToString());
            }
            catch (Exception ex)
            {
                Error err = new Error();
                err.LogError("错误信息", ex);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                listBox4.Items.Add("欠停周报开始：" + DateTime.Now.ToString());
                DataProcessing.zbaobiao16();
                listBox4.Items.Add("欠停周报结束：" + DateTime.Now.ToString());
            }
            catch (Exception ex)
            {
                Error err = new Error();
                err.LogError("错误信息", ex);
            }
        }
    }
}
