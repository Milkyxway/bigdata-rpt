using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
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
            listBox1.Items.Add("人工日报开始：" + DateTime.Now.ToString());
            int flag = 0;
            do
            {
                try
                {
                    DataProcessing.baobiao13();
                    //DataProcessing.baobiao14();
                    DataProcessing.baobiao1();
                    DataProcessing.baobiao2();
                    //DataProcessing.baobiao3();
                    //DataProcessing.baobiao5();
                    //DataProcessing.baobiao6();
                    DataProcessing.baobiao7();
                    DataProcessing.baobiao17();
                    DataProcessing.baobiao18();
                    //DataProcessing.baobiao19();
                    //DataProcessing.baobiao20();
                    //DataProcessing.baobiao8();
                    //DataProcessing.baobiao11();
                    DataProcessing.baobiao25();
                    DataProcessing.baobiao26();
                    DataProcessing.baobiao27();
                    DataProcessing.baobiao4();
                    DataProcessing.baobiao16();
                    flag = 10;
                    listBox1.Items.Add("人工日报结束：" + DateTime.Now.ToString());
                }
                catch (Exception ex)
                {
                    if (ex.ToString().Contains("Access denied")) { Mouse.denglu(); listBox1.Items.Add("重连" + DateTime.Now.ToString()); }
                    else { listBox1.Items.Add("人工日报失败：" + DateTime.Now.ToString()); }
                    Error err = new Error();err.LogError("错误信息", ex);
                    flag++;
                }
            }
            while (flag < 3);
        }
        private void button3_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    listBox1.Items.Add("人工周报开始：" + DateTime.Now.ToString());
            //    //DataProcessing.zbaobiao1();
            //    //DataProcessing.zbaobiao2();
            //    //DataProcessing.zbaobiao3();
            //    DataProcessing.zbaobiao4();
            //    DataProcessing.zbaobiao5();
            //    listBox1.Items.Add("人工周报结束：" + DateTime.Now.ToString());
            //}
            //catch (Exception ex)
            //{
            //    Error err = new Error();
            //    err.LogError("错误信息", ex);
            //}
        }
        private void button10_Click(object sender, EventArgs e)
        {
            listBox1.Items.Add("人工月报开始：" + DateTime.Now.ToString());
            int flag = 0;
            do
            {
                try
                {
                    DataProcessing.ybaobiao36();
                    DataProcessing.ybaobiao37();
                    DataProcessing.ybaobiao1();
                    DataProcessing.ybaobiao2();
                    DataProcessing.ybaobiao4();
                    DataProcessing.ybaobiao5();
                    DataProcessing.ybaobiao7();
                    DataProcessing.ybaobiao8();
                    DataProcessing.ybaobiao9();
                    DataProcessing.ybaobiao10();
                    DataProcessing.ybaobiao11();
                    DataProcessing.ybaobiao12();
                    DataProcessing.ybaobiao13();
                    DataProcessing.ybaobiao15();
                    DataProcessing.ybaobiao17();
                    DataProcessing.ybaobiao18();
                    DataProcessing.ybaobiao19();
                    DataProcessing.ybaobiao20();
                    DataProcessing.ybaobiao21();
                    DataProcessing.ybaobiao22();
                    DataProcessing.ybaobiao23();
                    DataProcessing.ybaobiao24();
                    DataProcessing.ybaobiao25();
                    DataProcessing.ybaobiao26();
                    DataProcessing.ybaobiao28();
                    DataProcessing.ybaobiao29();
                    DataProcessing.ybaobiao30();
                    DataProcessing.ybaobiao31();
                    DataProcessing.ybaobiao32();
                    DataProcessing.ybaobiao33();
                    DataProcessing.ybaobiao34();
                    DataProcessing.ybaobiao38();
                    DataProcessing.ybaobiao39();
                    DataProcessing.ybaobiao40();
                    DataProcessing.ybaobiao41();
                    flag = 10;
                    listBox1.Items.Add("人工月报结束：" + DateTime.Now.ToString());
                }
                catch (Exception ex)
                {
                    if (ex.ToString().Contains("Access denied")) { Mouse.denglu(); listBox1.Items.Add("失败重连" + DateTime.Now.ToString()); }
                    else { listBox1.Items.Add("人工月报失败：" + DateTime.Now.ToString()); }
                    Error err = new Error(); err.LogError("错误信息", ex);
                    flag++;
                }
            }
            while (flag < 3);
        }



        private void timer1_Tick(object sender, EventArgs e)
        {
            string dtt = DateTime.Now.ToString("HHmmss");// 得到 hour minute second  如果等于某个值就开始执行某个程序。

            if (dtt == "080000")//每天11:00:00 开始执行  071500  7:15:00
            {
                Mouse.denglu();
                listBox1.Items.Add("日报开始：" + DateTime.Now.ToString());
                int flag = 0;
                do
                {
                    try
                    {
                        DataProcessing.baobiao13();
                        //DataProcessing.baobiao14();
                        DataProcessing.baobiao1();
                        DataProcessing.baobiao2();
                        //DataProcessing.baobiao3();
                        //DataProcessing.baobiao5();
                        //DataProcessing.baobiao6();
                        DataProcessing.baobiao7();
                        DataProcessing.baobiao17();
                        DataProcessing.baobiao18();
                        //DataProcessing.baobiao19();
                        //DataProcessing.baobiao20();
                        //DataProcessing.baobiao8();
                        //DataProcessing.baobiao11();
                        ////DataProcessing.baobiao4();
                        ////DataProcessing.baobiao16();
                        DataProcessing.baobiao25();
                        DataProcessing.baobiao26();
                        DataProcessing.baobiao27();
                        flag = 10;
                        listBox1.Items.Add("日报结束：" + DateTime.Now.ToString());
                    }
                    catch (Exception ex)
                    {
                        if (ex.ToString().Contains("Access denied")) { Mouse.denglu(); listBox1.Items.Add("失败重连" + DateTime.Now.ToString()); }
                        else { listBox1.Items.Add("日报失败：" + DateTime.Now.ToString()); }
                        Error err = new Error(); err.LogError("错误信息", ex);
                        flag++;
                    }
                }
                while (flag < 3);
            }
            

            if (dtt == "120000")//每天12:00:00 开始执行  071500  7:15:00
            {
                Mouse.denglu();
                listBox1.Items.Add("日报宽表开始：" + DateTime.Now.ToString());
                int flag = 0;
                do
                {
                    try
                    {
                        //DataProcessing.baobiao13();
                        //DataProcessing.baobiao14();
                        //DataProcessing.baobiao1();
                        //DataProcessing.baobiao2();
                        //DataProcessing.baobiao3();
                        //DataProcessing.baobiao5();
                        //DataProcessing.baobiao6();
                        //DataProcessing.baobiao7();
                        //DataProcessing.baobiao17();
                        //DataProcessing.baobiao18();
                        //DataProcessing.baobiao19();
                        //DataProcessing.baobiao20();
                        //DataProcessing.baobiao8();
                        //DataProcessing.baobiao11();
                        DataProcessing.baobiao4();
                        DataProcessing.baobiao16();
                        flag = 10;
                        listBox1.Items.Add("日报宽表结束：" + DateTime.Now.ToString());
                    }
                    catch (Exception ex)
                    {
                        if (ex.ToString().Contains("Access denied")) { Mouse.denglu(); listBox1.Items.Add("失败重连" + DateTime.Now.ToString()); }
                        else { listBox1.Items.Add("日报宽表失败：" + DateTime.Now.ToString()); }
                        Error err = new Error(); err.LogError("错误信息", ex);
                        flag++;
                    }
                }
                while (flag < 3);

                listBox1.Items.Add("工程宽表日报开始：" + DateTime.Now.ToString());
                flag = 0;
                do
                {
                    try
                    {
                        DataProcessing.gckb();
                        DataProcessing.gc1();
                        DataProcessing.gc2();
                        DataProcessing.gc3();
                        DataProcessing.gc4();
                        DataProcessing.gc5();
                        DataProcessing.gc6();
                        DataProcessing.gc7();
                        flag = 10;
                        listBox1.Items.Add("工程宽表日报结束：" + DateTime.Now.ToString());
                    }
                    catch (Exception ex)
                    {
                        if (ex.ToString().Contains("Access denied")) { Mouse.denglu(); listBox1.Items.Add("失败重连" + DateTime.Now.ToString()); }
                        else { listBox1.Items.Add("工程宽表失败：" + DateTime.Now.ToString()); }
                        Error err = new Error(); err.LogError("错误信息", ex);
                        flag++;
                    }
                }
                while (flag < 3);

                if (DateTime.Now.ToString("dd") == "03") //每月三号执行月报
                {
                    listBox1.Items.Add("月报开始：" + DateTime.Now.ToString());
                    flag = 0;
                    do
                    {
                        try
                        {
                            DataProcessing.ybaobiao36();
                            DataProcessing.ybaobiao37();
                            DataProcessing.ybaobiao1();
                            DataProcessing.ybaobiao2();
                            DataProcessing.ybaobiao4();
                            DataProcessing.ybaobiao5();
                            DataProcessing.ybaobiao7();
                            DataProcessing.ybaobiao8();
                            DataProcessing.ybaobiao9();
                            DataProcessing.ybaobiao10();
                            DataProcessing.ybaobiao11();
                            DataProcessing.ybaobiao12();
                            DataProcessing.ybaobiao13();
                            DataProcessing.ybaobiao15();
                            DataProcessing.ybaobiao17();
                            DataProcessing.ybaobiao18();
                            DataProcessing.ybaobiao19();
                            DataProcessing.ybaobiao20();
                            DataProcessing.ybaobiao21();
                            DataProcessing.ybaobiao22();
                            DataProcessing.ybaobiao23();
                            DataProcessing.ybaobiao24();
                            DataProcessing.ybaobiao25();
                            DataProcessing.ybaobiao26();
                            DataProcessing.ybaobiao28();
                            DataProcessing.ybaobiao29();
                            DataProcessing.ybaobiao30();
                            DataProcessing.ybaobiao31();
                            DataProcessing.ybaobiao32();
                            DataProcessing.ybaobiao33();
                            DataProcessing.ybaobiao34();
                            DataProcessing.ybaobiao38();
                            DataProcessing.ybaobiao39();
                            DataProcessing.ybaobiao40();
                            DataProcessing.ybaobiao41();
                            flag = 10;
                            listBox1.Items.Add("月报结束：" + DateTime.Now.ToString());
                        }
                        catch (Exception ex)
                        {
                            if (ex.ToString().Contains("Access denied")) { Mouse.denglu(); listBox1.Items.Add("失败重连" + DateTime.Now.ToString()); }
                            else { listBox1.Items.Add("月报失败：" + DateTime.Now.ToString()); }
                            Error err = new Error(); err.LogError("错误信息", ex);
                            flag++;
                        }
                    }
                    while (flag < 3);
                }

                if (DateTime.Now.ToString("dd") == "26") //每月26号执行江阴互动数月报
                {
                    listBox1.Items.Add("江阴互动数开始：" + DateTime.Now.ToString());
                    flag = 0;
                    do
                    {
                        try
                        {
                            DataProcessing.ybaobiao12();
                            flag = 10;
                            listBox1.Items.Add("江阴互动数结束：" + DateTime.Now.ToString());
                        }
                        catch (Exception ex)
                        {
                            if (ex.ToString().Contains("Access denied")) { Mouse.denglu(); listBox1.Items.Add("失败重连" + DateTime.Now.ToString()); }
                            else { listBox1.Items.Add("江阴互动数失败：" + DateTime.Now.ToString()); }
                            Error err = new Error(); err.LogError("错误信息", ex);
                            flag++;
                        }
                    }
                    while (flag < 3);
                }


                if (DateTime.Now.ToString("dd") == "1" || DateTime.Now.ToString("dd") == "11" || DateTime.Now.ToString("dd") == "21") //每月1\11\21号执行江阴互动数月报
                {
                    listBox1.Items.Add("客服欠费未复通：" + DateTime.Now.ToString());
                    flag = 0;
                    do
                    {
                        try
                        {
                            DataProcessing.baobiao9();
                            flag = 10;
                            listBox1.Items.Add("客服欠费未复通：" + DateTime.Now.ToString());
                        }
                        catch (Exception ex)
                        {
                            if (ex.ToString().Contains("Access denied")) { Mouse.denglu(); listBox1.Items.Add("失败重连" + DateTime.Now.ToString()); }
                            else { listBox1.Items.Add("客服欠费未复通失败：" + DateTime.Now.ToString()); }
                            Error err = new Error(); err.LogError("错误信息", ex);
                            flag++;
                        }
                    }
                    while (flag < 3);
                }


                if (DateTime.Now.ToString("dd") == "15") //每月15号执行各业务新增欠费金额统计
                {
                    listBox1.Items.Add("月新增欠费金额开始：" + DateTime.Now.ToString());
                    flag = 0;
                    do
                    {
                        try
                        {
                            DataProcessing.ybaobiao3();
                            flag = 10;
                            listBox1.Items.Add("月新增欠费金额结束：" + DateTime.Now.ToString());

                        }
                        catch (Exception ex)
                        {
                            if (ex.ToString().Contains("Access denied")) { Mouse.denglu(); listBox1.Items.Add("失败重连" + DateTime.Now.ToString()); }
                            else { listBox1.Items.Add("月新增欠费金额失败：" + DateTime.Now.ToString()); }
                            Error err = new Error(); err.LogError("错误信息", ex);
                            flag++;
                        }
                    }
                    while (flag < 3);
                }


                //try
                //{
                //    if (DateTime.Now.DayOfWeek.ToString() == "Wednesday")
                //    {
                //        listBox2.Items.Add("欠停周报开始：" + DateTime.Now.ToString());
                //        //DataProcessing.zbaobiao11();
                //        listBox2.Items.Add("欠停周报结束：" + DateTime.Now.ToString());
                //        label8.Text = (Int32.Parse(label8.Text) + 1).ToString();
                //        label18.Text = (Int32.Parse(label18.Text) + 1).ToString();
                //        label10.Text = "1";
                //    }
                //}
                //catch (Exception ex)
                //{
                //    label10.Text = "0";
                //    label13.Text = (Int32.Parse(label13.Text) + 1).ToString();
                //    Error err = new Error();
                //    err.LogError("错误信息", ex);
                //}
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            listBox1.Items.Add("人工月新增欠费金额开始：" + DateTime.Now.ToString());
            int flag = 0;
            do
            {
                try
                {
                    DataProcessing.ybaobiao3();
                    flag = 10;
                    listBox1.Items.Add("人工月新增欠费金额结束：" + DateTime.Now.ToString());

                }
                catch (Exception ex)
                {
                    if (ex.ToString().Contains("Access denied")) { Mouse.denglu(); listBox1.Items.Add("失败重连" + DateTime.Now.ToString()); }
                    else { listBox1.Items.Add("人工月新增欠费金额失败：" + DateTime.Now.ToString()); }
                    Error err = new Error(); err.LogError("错误信息", ex);
                    flag++;
                }
            }
            while (flag < 3);
        }

     

        private void button7_Click(object sender, EventArgs e)
        {
            listBox1.Items.Add("26号互动数开始：" + DateTime.Now.ToString());
            int flag = 0;
            do
            {
                try
                {
                    DataProcessing.ybaobiao12();
                    flag = 10;
                    listBox1.Items.Add("26号互动数结束：" + DateTime.Now.ToString());
                }
                catch (Exception ex)
                {
                    if (ex.ToString().Contains("Access denied")) { Mouse.denglu(); listBox1.Items.Add("失败重连" + DateTime.Now.ToString()); }
                    else { listBox1.Items.Add("26号互动数失败：" + DateTime.Now.ToString()); }
                    Error err = new Error(); err.LogError("错误信息", ex);
                    flag++;
                }
            }
            while (flag < 3);
        }



        private void button6_Click(object sender, EventArgs e)
        {
            string sqlString =
"select rad.jy_region_name 区域,\n" +
"      count( distinct cust.cust_code) 缴费客户数\n" +
"  from cp2.cm_customer cust，\n" +
"       rep2.rep_fact_um_subscriber_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddDays(-4).ToString("yyyyMMdd") + " subs,\n" +
"       rep2.rep_fact_cust_info_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddDays(-4).ToString("yyyyMMdd") + "     info,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       files2.cm_account        acct\n" +
" where cust.corp_org_id = 3328\n" +
"   and cust.cust_id = subs.cust_id    and  rad.cust_id=cust.cust_id\n" +
"   and cust.cust_id = acct.cust_id\n" +
"  and cust.cust_id = info.cust_id\n" +
" and info.iden_addr_name like '%" + textBox1.Text + "%'\n" +
"group by   rad.jy_region_name";
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    listBox1.Items.Add(dt.Rows[i]["区域"].ToString() + dt.Rows[i]["缴费客户数"].ToString() + textBox1.Text);
                }
            }
            else
            {
                listBox1.Items.Add("未查询到" + textBox1.Text);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            listBox1.Items.Add("人工工程宽表开始：" + DateTime.Now.ToString());
            int flag = 0;
            do
            {
                try
                {
                    DataProcessing.gckb();
                    flag = 10;
                    listBox1.Items.Add("人工工程宽表结束：" + DateTime.Now.ToString());
                }
                catch (Exception ex)
                {
                    if (ex.ToString().Contains("Access denied")) { Mouse.denglu(); listBox1.Items.Add("失败重连" + DateTime.Now.ToString()); }
                    else { listBox1.Items.Add("人工工程宽表失败：" + DateTime.Now.ToString()); }
                    Error err = new Error(); err.LogError("错误信息", ex);
                    flag++;
                }
            }
            while (flag < 3);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            listBox1.Items.Add("批量工程校验开始" + DateTime.Now.ToString());
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            DataTable dt = ExcelHelper.InputFromExcel(pathInfo.Parent.FullName + "\\工程源文件\\批量源.xls", "Sheet1");

            StringBuilder st1 = new StringBuilder();
            StringBuilder st2 = new StringBuilder();

           
            if (dt.Rows.Count > 0)
            {
                st1.Append("when addr.std_addr_name like '%" + dt.Rows[0]["小区"].ToString() + "%' then '" + dt.Rows[0]["id"].ToString() + "," + dt.Rows[0]["小区"].ToString() + "'\n");
                st2.Append("( addr.std_addr_name like '%" + dt.Rows[0]["区域"].ToString() + "%' and  addr.std_addr_name like '%" + dt.Rows[0]["小区"].ToString() + "%')\n");
                for (int i = 1; i < dt.Rows.Count; i++)
                {
                    string[] ss = dt.Rows[i]["小区"].ToString().Split(',');

                    if (ss.Length == 1)
                    {

                        st1.Append("when addr.std_addr_name like '%" + dt.Rows[i]["小区"].ToString() + "%' then '" + dt.Rows[i]["id"].ToString() + "," + dt.Rows[i]["小区"].ToString() + "'\n");
                        st2.Append("or (addr.std_addr_name like '%" + dt.Rows[i]["区域"].ToString() + "%' and  addr.std_addr_name like '%" + dt.Rows[i]["小区"].ToString() + "%')\n");
                    }
                    else
                    {
                        st1.Append("when （addr.std_addr_name like '%" + ss[0] + "%' ");
                        st2.Append("or (addr.std_addr_name like '%" + dt.Rows[i]["区域"].ToString() + "%' and  （addr.std_addr_name like '%" + ss[0] + "%' ");
                        for (int j = 1; j < ss.Length; j++)
                        {
                            st1.Append(" or addr.std_addr_name like '%" + ss[j] + "%'");
                            st2.Append(" or addr.std_addr_name like '%" + ss[j] + "%'");
                        }
                        st1.Append(") then '" + dt.Rows[i]["id"].ToString() + "," + dt.Rows[i]["小区"].ToString() + "'\n");
                        st2.Append("))\n");
                    }

                }
            }
            
            string sqlString =
"select a.*,\n" +
"       case\n" +
"         when exists\n" +
"          (select 1\n" +
"                 from (select distinct cust.cust_id,\n" +
"                                       case\n" +
st1.ToString() +
"                                         else\n" +
"                                          ''\n" +
"                                       end 小区,\n" +
"                                       addr.std_addr_name 地址\n" +
"                         from cp2.cm_customer cust, files2.um_address addr\n" +
"                        where cust.cust_id = addr.cust_id\n" +
"                          and cust.own_corp_org_id = 3328\n" +
"                          and (\n" +
st2.ToString() +
"                              )\n" +
"                          and exists\n" +
"                        (select 1\n" +
"                                 from files2.um_subscriber   subs,\n" +
"                                      files2.um_offer_06     ofer,\n" +
"                                      files2.um_offer_sta_02 fsta\n" +
"                                where subs.subscriber_ins_id =\n" +
"                                      ofer.subscriber_ins_id\n" +
"                                  and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"                                  and fsta.expire_date > sysdate\n" +
"                                  and ofer.expire_date > sysdate\n" +
"                                  and ofer.prod_service_id in (1002, 1004)\n" +
"                                  and fsta.os_status is null\n" +
"                                  and subs.cust_id = cust.cust_id)) b\n" +
"                where b.小区 = a.小区) then\n" +
"          '有'\n" +
"         else\n" +
"          ''\n" +
"       end 是否有缴费客户数\n" +
"  from (select distinct rad.jy_region_name 区域, tmp.std_addr_name 小区\n" +
"          from wxjy.jy_region_address_rel rad,\n" +
"               (select addr.cust_addr_id,\n" +
"                       case\n" +
st1.ToString() +
"                         else\n" +
"                          ''\n" +
"                       end std_addr_name\n" +
"                  from files2.um_address addr\n" +
"                 where (\n" +
st2.ToString() +
"                              )\n" +
"                   and addr.own_corp_org_id = 3328) tmp\n" +
"         where rad.cust_addr_id = tmp.cust_addr_id) a";

            DataTable dataTable = OracleHelper.ExecuteDataTable(sqlString);

            Excel.Application xapp = new Excel.Application();
            string filepath = pathInfo.Parent.FullName + "\\工程源文件\\批量源.xls";
            Excel.Workbook xbook = xapp.Workbooks._Open(filepath, Missing.Value, Missing.Value,
                                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            Excel.Worksheet xsheet = (Excel.Worksheet)xbook.Sheets[1];


            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                string[] sArray = dataTable.Rows[i]["小区"].ToString().Split(',');
                foreach (string j in sArray)
                {
                    Regex rex = new Regex(@"^\d+$");
                    if (rex.IsMatch(j.ToString()))
                    {
                        Excel.Range rng2 = xsheet.get_Range("F" + j.ToString(), Missing.Value);
                        rng2.Value2 = dataTable.Rows[i]["是否有缴费客户数"].ToString();
                    }
                }
            }

            Excel.Worksheet xsheet2 = (Excel.Worksheet)xbook.Sheets[2];
            Excel.Range rng1 = xsheet2.get_Range("A1", Missing.Value);
            rng1.Value2 = sqlString;
            xbook.SaveAs(pathInfo.Parent.FullName + "\\工程\\批量校验结果" + DateTime.Now.ToString("yyyyMMdd") + ".xls", Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            xsheet = null;
            xbook = null;
            xapp.Quit();
            xapp = null;
            GC.Collect();
            listBox1.Items.Add("批量工程校验结束" + DateTime.Now.ToString());
        }

        private void button11_Click(object sender, EventArgs e)
        {
            listBox1.Items.Add("工程明细开始：" + DateTime.Now.ToString());
            int flag = 0;
            do
            {
                try
                {
                    DataProcessing.gc1();
                    DataProcessing.gc2();
                    DataProcessing.gc3();
                    DataProcessing.gc4();
                    DataProcessing.gc5();
                    DataProcessing.gc6();
                    DataProcessing.gc7();
                    flag = 10;
                    listBox1.Items.Add("工程明细结束：" + DateTime.Now.ToString());
                }
                catch (Exception ex)
                {
                    if (ex.ToString().Contains("Access denied")) { Mouse.denglu(); listBox1.Items.Add("失败重连" + DateTime.Now.ToString()); }
                    else { listBox1.Items.Add("工程明细失败：" + DateTime.Now.ToString()); }
                    Error err = new Error(); err.LogError("错误信息", ex);
                    flag++;
                }
            }
            while (flag < 3);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            //listBox1.Items.Add("人工月报开始：" + DateTime.Now.ToString());
            //int flag = 0;
            //do
            //{
            //    try
            //    {

            //        DataProcessing.ybaobiao40();
            //        flag = 10;
            //        listBox1.Items.Add("人工月报结束：" + DateTime.Now.ToString());
            //    }
            //    catch (Exception ex)
            //    {
            //        if (ex.ToString().Contains("Access denied")) { Mouse.denglu(); listBox1.Items.Add("失败重连" + DateTime.Now.ToString()); }
            //        else { listBox1.Items.Add("人工月报失败：" + DateTime.Now.ToString()); }
            //        Error err = new Error(); err.LogError("错误信息", ex);
            //        flag++;
            //    }
            //}
            //while (flag < 3);

            listBox1.Items.Add("人工日报开始：" + DateTime.Now.ToString());
            int flag = 0;
            do
            {
                try
                {
                    //DataProcessing.baobiao11();
                    //DataProcessing.baobiao13();
                    //DataProcessing.baobiao14();
                    //DataProcessing.baobiao1();
                    //DataProcessing.baobiao2();
                    //DataProcessing.baobiao3();
                    //DataProcessing.baobiao5();
                    //DataProcessing.baobiao6();
                    //DataProcessing.baobiao7();
                    //DataProcessing.baobiao17();
                    //DataProcessing.baobiao18();
                    //DataProcessing.baobiao8();
                    //DataProcessing.baobiao19();
                    //DataProcessing.baobiao20();
                    //DataProcessing.baobiao4();
                    //DataProcessing.baobiao16();
                    DataProcessing.baobiao25();
                    //DataProcessing.baobiao26();
                    flag = 10;
                    listBox1.Items.Add("人工日报结束：" + DateTime.Now.ToString());
                }
                catch (Exception ex)
                {
                    if (ex.ToString().Contains("Access denied"))
                    { Mouse.denglu(); listBox1.Items.Add("重连" + DateTime.Now.ToString()); }
                    Error err = new Error();
                    err.LogError("错误信息", ex);
                    flag++;
                }
            }
            while (flag < 3);

            //flag = 0;
            //do
            //{
            //    try
            //    {
            //        //DataProcessing.baobiao11();
            //        //DataProcessing.baobiao13();
            //        //DataProcessing.baobiao14();
            //        //DataProcessing.baobiao1();
            //        DataProcessing.baobiao2();
            //        //DataProcessing.baobiao3();
            //        //DataProcessing.baobiao5();
            //        //DataProcessing.baobiao6();
            //        //DataProcessing.baobiao7();
            //        //DataProcessing.baobiao17();
            //        //DataProcessing.baobiao18();
            //        //DataProcessing.baobiao8();
            //        //DataProcessing.baobiao19();
            //        //DataProcessing.baobiao20();
            //        //DataProcessing.baobiao4();
            //        //DataProcessing.baobiao16();
            //        flag = 10;
            //        listBox1.Items.Add("人工日报结束：" + DateTime.Now.ToString());
            //    }
            //    catch (Exception ex)
            //    {
            //        if (ex.ToString().Contains("Access denied"))
            //        { Mouse.denglu(); listBox1.Items.Add("重连" + DateTime.Now.ToString()); }
            //        Error err = new Error();
            //        err.LogError("错误信息", ex);
            //        flag++;
            //    }
            //}
            //while (flag < 3);



    //        string sqlString =

    //"select rad.jy_region_name 区域,\n" +
    //"       tmp.std_addr_name 小区,\n" +
    //"     count(distinct case\n" +
    //"           when ofer.prod_service_id = 1002 and not exists\n" +
    //"            (select 1\n" +
    //"                   from files2.um_subscriber subs1,\n" +
    //"                        files2.um_res        ures1,\n" +
    //"                        res1.res_terminal    term1,\n" +
    //"                        res1.res_sku         rsku1\n" +
    //"                  where subs1.subscriber_ins_id = ures1.subscriber_ins_id\n" +
    //"                    and ures1.res_equ_no = term1.serial_no\n" +
    //"                    and term1.res_sku_id = rsku1.res_sku_id\n" +
    //"                    and ures1.res_type_id = 2\n" +
    //"                    and rsku1.res_sku_name in\n" +
    //"                        ('银河高清基本型HDC6910(江阴)',\n" +
    //"                         '银河高清交互型HDC691033(江阴)',\n" +
    //"                         '银河智能高清交互型HDC6910798(江阴)',\n" +
    //"                         '银河4K交互型HDC691090',\n" +
    //"                         '4K超高清型II型融合型（EOC）',\n" +
    //"                         '银河4K超高清II型融合型HDC6910B1(江阴)',\n" +
    //"                         '4K超高清简易型（基本型）')\n" +
    //"                    and subs1.cust_id = cust.cust_id) then\n" +
    //"            cust.cust_id\n" +
    //"           else\n" +
    //"            null\n" +
    //"         end) 纯标清客户缴费数,\n" +
    //"       count(distinct case\n" +
    //"               when ofer.prod_service_id = 1004 and exists\n" +
    //"                (select 1\n" +
    //"                       from upc1.pm_offer prod\n" +
    //"                      where (prod.offer_name like '%300M%' or prod.offer_name like '%500M%' or  prod.offer_name like '%600M%'\n" +
    //"                           or prod.offer_name like '%1000M%')\n" +
    //"                        and prod.offer_id = fsta.offer_id) then\n" +
    //"                cust.cust_id\n" +
    //"               else\n" +
    //"                null\n" +
    //"             end) 缴费客户数FTTH按产品\n" +
    //"  from cp2.cm_customer cust,\n" +
    //"       files2.um_subscriber subs,\n" +
    //"       files2.um_offer_06 ofer,\n" +
    //"       files2.um_offer_sta_02 fsta,\n" +
    //"       wxjy.jy_region_address_rel rad,\n" +
    //"       wxjy.jy_customer_ftth_rel tmp\n" +
    //" where cust.cust_id = rad.cust_id\n" +
    //"  and  cust.cust_id = tmp.cust_id\n" +
    //"   and cust.cust_id = subs.cust_id\n" +
    //"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
    //"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
    //"   and fsta.expire_date > sysdate\n" +
    //"   and ofer.expire_date > sysdate\n" +
    //"   and ofer.prod_service_id in (1002, 1004)\n" +
    //"   and fsta.offer_status = '1'\n" +
    //"   and fsta.os_status is null\n" +
    //"   and cust.own_corp_org_id = 3328\n" +
    //"   and tmp.std_addr_name is not null\n" +
    //" group by rad.jy_region_name, tmp.std_addr_name";

    //        textBox2.Text = sqlString;


            //string str = "水芝苑 ,1-40幢;水芝苑 15栋;夏港元 ,2-60幢";


            ////  textBox2.Text = str.Substring(0, 10) + "ceshi" + str.Substring(10, str.Length - 10);


            //string[] ff = str.Split(';');//幢范围加独立栋:天鹤,1-5幢d天鹤6栋,天鹤10栋,天鹤22栋

            //if(ff.Length > 1)
            //{ 

            //    string[] ss = ff[0].Split(',');

            //    for (int j = 0; j < ss.Length; j++)
            //    {
            //        listBox1.Items.Add(ss[j]);

            //    }

            ////    string[] dd = str.Split('d');//幢范围加独立栋:天鹤,1-5幢d天鹤6栋,天鹤10栋,天鹤22栋
            ////if (dd.Length > 1)//判断幢范围加独立栋
            ////{
            ////    string[] ss = dd[0].Split(',');
            ////    Regex reg = new Regex(@"[0-9]+");//查出所有数字
            ////    textBox2.Text = dd[0] + "/n" + dd[1];


            ////    //for (int k = int.Parse(mc[0].Value) + 1; k < int.Parse(mc[1].Value) + 1; k++)
            ////    //{
            ////    //    st1.Append(" or addr.std_addr_name like '%" + ss[0] + k.ToString() + "幢%'");
            ////    //    st2.Append(" or addr.std_addr_name like '%" + ss[0] + k.ToString() + "幢%'");
            ////    //}
            ////    string[] ss1 = dd[1].Split(',');
            ////    for (int j = 0; j < ss.Length; j++)
            ////    {
            ////        listBox1.Items.Add(ss1[j]);

            ////    }

            ////}

            //}


            //if (dd.Length > 1)//判断幢范围加独立栋
            //{
            //    string[] ss = dd[0].Split(',');
            //    Regex reg = new Regex(@"[0-9]+");//查出所有数字
            //    MatchCollection mc = reg.Matches(str);

            //    for (int k = int.Parse(mc[0].Value) + 1; k < int.Parse(mc[1].Value) + 1; k++)
            //    {
            //        listBox1.Items.Add(" or addr.std_addr_name like '%" + ss[0] + k.ToString() + "幢%'");
            //    }
            //    string[] ss1 = dd[1].Split(',');
            //    for (int j = 0; j < ss.Length; j++)
            //    {
            //        listBox1.Items.Add(ss1[j]);

            //    }

            //}

            //string[] dd = str.Split('d');//幢范围加独立栋:天鹤,1-5幢d天鹤6栋,天鹤10栋,天鹤22栋
            //string[] uu = str.Split('u');//幢范围加栋范围:天鹤,1-5幢u6-10栋
            //string[] bb = str.Split('-');//幢范围:天鹤,1-5幢

            //if (dd.Length > 1)//判断幢范围加独立栋
            //{
            //    listBox1.Items.Add("幢范围加独立栋");
            //}
            //else if (uu.Length > 1)//判断幢范围加栋范围
            //{
            //    listBox1.Items.Add("幢范围加栋范围");
            //}
            //else if (bb.Length > 1)//判断幢范围
            //{
            //    listBox1.Items.Add("幢范围");
            //}
            //else//光小区
            //{
            //    listBox1.Items.Add("光小区");
            //}



            ////     string[] dd = str.Split('d');
            //     string[] ss = str.Split('-');
            //     //Regex reg = new Regex(@"[0-9]+");//查出所有数字
            //     //MatchCollection mc = reg.Matches(dt.Rows[i]["小区"].ToString());
            //     Regex reg = new Regex(@"[0-9]+");//查出所有数字
            //     MatchCollection mc = reg.Matches(str);

            //     listBox1.Items.Add(ss[0]);
            //     listBox1.Items.Add(ss[1]);
            //     listBox1.Items.Add(ss[2]);
            //     listBox1.Items.Add(mc[0].Value);
            //     listBox1.Items.Add(mc[1].Value);
            //     listBox1.Items.Add(mc[2].Value);
            //     listBox1.Items.Add(mc[3].Value);

        }



        private  void button13_Click(object sender, EventArgs e)
        {
            Mouse.denglu();


            //// 是鼠标自动到（100，100）位置
            //Mouse.MouseMoveToPoint(330, 400);
            //Mouse.WaitFunctions(1);
            //Mouse.LeftClick();
            //Mouse.WaitFunctions(1);
            //Mouse.MouseMoveToPoint(300, 450);
            //Mouse.WaitFunctions(1);
            //Mouse.LeftClick();
            //Mouse.WaitFunctions(1);
            //Mouse.MouseMoveToPoint(100, 215);
            //Mouse.WaitFunctions(1);
            //Mouse.LeftClick();
            //Mouse.WaitFunctions(1);
            //Mouse.KeyboardInput("jy_szw");
            //Mouse.WaitFunctions(1);
            //Mouse.MouseMoveToPoint(100, 315);
            //Mouse.WaitFunctions(1);
            //Mouse.LeftClick();
            //Mouse.WaitFunctions(1);
            //Mouse.KeyboardInput("Aa@6543211");
            //Mouse.WaitFunctions(1);
            //Mouse.MouseMoveToPoint(100, 380);
            //Mouse.WaitFunctions(1);
            //Mouse.LeftClick();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            listBox1.Items.Add("创建临时表开始：" + DateTime.Now.ToString());
            if (DataProcessing.gckbst())
            {
                listBox1.Items.Add("创建临时表成功：" + DateTime.Now.ToString());
            }
            else
            {
                listBox1.Items.Add("创建临时表失败：" + DateTime.Now.ToString());
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            listBox1.Items.Add("删除临时表开始：" + DateTime.Now.ToString());
            if (DataProcessing.gckbdel())
            {
                listBox1.Items.Add("删除临时表成功：" + DateTime.Now.ToString());
            }
            else
            {
                listBox1.Items.Add("删除临时表失败：" + DateTime.Now.ToString());
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            listBox1.Items.Add("FTTH四套开始：" + DateTime.Now.ToString());
            int flag = 0;
            do
            {
                try
                {
                    DataProcessing.baobiao21();
                    DataProcessing.baobiao22();
                    DataProcessing.baobiao23();
                    DataProcessing.baobiao24();
                    flag = 10;
                    listBox1.Items.Add("FTTH四套结束：" + DateTime.Now.ToString());
                }
                catch (Exception ex)
                {
                    if (ex.ToString().Contains("Access denied")) { Mouse.denglu(); listBox1.Items.Add("重连" + DateTime.Now.ToString()); }
                    else { listBox1.Items.Add("FTTH四套失败：" + DateTime.Now.ToString()); }
                    Error err = new Error(); err.LogError("错误信息", ex);
                    flag++;
                }
            }
            while (flag < 3);
        }
    }
}
