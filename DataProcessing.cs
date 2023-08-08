using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace AutoUpDataBoss
{
    class DataProcessing
    {
        #region 删除FTTH临时表
        public static bool gckbdel()
        {
            return OracleHelper.ExecuteNonQuery("drop table wxjy.jy_customer_ftth_rel");
        }
        #endregion

        #region 创建FTTH所需临时表
        public static bool gckbst()
        {
            DirectoryInfo pathInfo1 = new DirectoryInfo(Environment.CurrentDirectory);
            DataTable dt = ExcelHelper.InputFromExcel(pathInfo1.Parent.FullName + "\\工程源文件\\小区源.xls", "Sheet1");
            StringBuilder st1 = new StringBuilder();
            if (dt.Rows.Count > 0)
            {
                st1.Append("when (rad.jy_region_name = '" + dt.Rows[0]["区域"].ToString() + "' and addr.std_addr_name like '%" + dt.Rows[0]["小区"].ToString() + "%') then '" + dt.Rows[0]["ID"].ToString() + "," + dt.Rows[0]["小区"].ToString() + "'\n");
                for (int i = 1; i < dt.Rows.Count; i++)
                {
                    string[] ff = dt.Rows[i]["小区"].ToString().Split('!');
                    if (ff.Length > 1)
                    {
                        st1.Append("when (rad.jy_region_name = '" + dt.Rows[i]["区域"].ToString() + "' and (addr.std_addr_name like '无无' ");
                        for (int a = 0; a < ff.Length; a++)
                        {
                            string[] aa = ff[a].Split('a');//通用范围性地址:天鹤,1-5,幢
                            string[] bb = ff[a].Split('b');//单数范围:益健路,15-45,号b
                            string[] cc = ff[a].Split('c');//双数范围:人民路,210-362,幢c
                                                           //string[] zz = ff[a].Split('z');//幢范围加独立栋:天鹤,1-5幢z
                                                           //string[] dd = ff[a].Split('d');//幢范围加独立栋:天鹤,1-5栋d
                                                           //string[] hh = ff[a].Split('h');//幢范围加独立号:天鹤,1-5号h
                            string[] xq = ff[a].Split('q');//光小区
                                                           //string[] uu = ff[a].Split('u');//幢范围加栋范围:天鹤,1-5幢u6-10栋
                                                           //string[] bb = ff[a].Split('-');//幢范围:天鹤,1-5幢
                            if (aa.Length > 1)//通用范围性地址
                            {
                                string[] ss = aa[0].Split(',');
                                Regex reg = new Regex(@"[0-9]+");//查出所有数字
                                MatchCollection mc = reg.Matches(ss[1]);
                                for (int k = int.Parse(mc[0].Value); k < int.Parse(mc[1].Value) + 1; k++)
                                {
                                    st1.Append(" or addr.std_addr_name like '%" + ss[0] + k.ToString() + ss[2] + "%'");
                                }
                            }
                            else if (bb.Length > 1)//通用范围性地址  单数
                            {
                                string[] ss = bb[0].Split(',');
                                Regex reg = new Regex(@"[0-9]+");//查出所有数字
                                MatchCollection mc = reg.Matches(ss[1]);
                                for (int k = int.Parse(mc[0].Value); k < int.Parse(mc[1].Value) + 1; k++)
                                {
                                    if (k % 2 != 0)//判断是否单数
                                    {
                                        st1.Append(" or addr.std_addr_name like '%" + ss[0] + k.ToString() + ss[2] + "%'");
                                    }
                                }
                            }
                            else if (cc.Length > 1)//通用范围性地址  双数
                            {
                                string[] ss = cc[0].Split(',');
                                Regex reg = new Regex(@"[0-9]+");//查出所有数字
                                MatchCollection mc = reg.Matches(ss[1]);
                                for (int k = int.Parse(mc[0].Value); k < int.Parse(mc[1].Value) + 1; k++)
                                {
                                    if (k % 2 == 0)//判断是否偶数
                                    {
                                        st1.Append(" or addr.std_addr_name like '%" + ss[0] + k.ToString() + ss[2] + "%'");
                                    }
                                }
                            }
                            //if (zz.Length > 1)//判断幢范围
                            //{
                            //    string[] ss = zz[0].Split(',');
                            //    Regex reg = new Regex(@"[0-9]+");//查出所有数字
                            //    MatchCollection mc = reg.Matches(ss[1]);
                            //    for (int k = int.Parse(mc[0].Value); k < int.Parse(mc[1].Value) + 1; k++)
                            //    {
                            //        st1.Append(" or addr.std_addr_name like '%" + ss[0] + k.ToString() + "幢%'");
                            //    }
                            //}
                            //else if (dd.Length > 1)//判断栋范围
                            //{
                            //    string[] ss = dd[0].Split(',');
                            //    Regex reg = new Regex(@"[0-9]+");//查出所有数字
                            //    MatchCollection mc = reg.Matches(ss[1]);
                            //    for (int k = int.Parse(mc[0].Value); k < int.Parse(mc[1].Value) + 1; k++)
                            //    {
                            //        st1.Append(" or addr.std_addr_name like '%" + ss[0] + k.ToString() + "栋%'");
                            //    }
                            //}
                            //else if (hh.Length > 1)//判断号范围
                            //{
                            //    string[] ss = hh[0].Split(',');
                            //    Regex reg = new Regex(@"[0-9]+");//查出所有数字
                            //    MatchCollection mc = reg.Matches(ss[1]);
                            //    for (int k = int.Parse(mc[0].Value); k < int.Parse(mc[1].Value) + 1; k++)
                            //    {
                            //        st1.Append(" or addr.std_addr_name like '%" + ss[0] + k.ToString() + "号%'");
                            //    }
                            //}


                            else if (xq.Length > 1)//纯小区
                            {
                                string[] ss = xq[0].Split(',');
                                if (ss.Length == 1)
                                {

                                    st1.Append(" or addr.std_addr_name like '%" + ss[0] + "%'");
                                }
                                else
                                {
                                    for (int j = 0; j < ss.Length; j++)
                                    {
                                        st1.Append(" or addr.std_addr_name like '%" + ss[j] + "%'");
                                    }
                                }
                            }
                        }
                        st1.Append(")) then '" + dt.Rows[i]["ID"].ToString() + "," + dt.Rows[i]["小区"].ToString() + "'\n");
                    }
                    else//光小区
                    {
                        string[] ss = dt.Rows[i]["小区"].ToString().Split(',');
                        if (ss.Length == 1)
                        {

                            st1.Append("when (rad.jy_region_name = '" + dt.Rows[i]["区域"].ToString() + "' and addr.std_addr_name like '%" + dt.Rows[i]["小区"].ToString() + "%') then '" + dt.Rows[i]["ID"].ToString() + "," + dt.Rows[i]["小区"].ToString() + "'\n");
                        }
                        else
                        {
                            st1.Append("when (rad.jy_region_name = '" + dt.Rows[i]["区域"].ToString() + "' and (addr.std_addr_name like '%" + ss[0] + "%'");

                            for (int j = 1; j < ss.Length; j++)
                            {
                                st1.Append(" or addr.std_addr_name like '%" + ss[j] + "%'");
                            }
                            st1.Append(")) then '" + dt.Rows[i]["ID"].ToString() + "," + dt.Rows[i]["小区"].ToString() + "'\n");
                        }
                    }
                }
            }
            string sqlString =
"create table wxjy.jy_customer_ftth_rel as\n" +
"select * from (select addr.cust_id,rad.jy_region_name,\n" +
"               case\n" +
st1.ToString() +
"                 else\n" +
"                  ''\n" +
"               end std_addr_name\n" +
"          from  wxjy.jy_region_address_rel rad,files2.um_address addr\n" +
"         where  addr.own_corp_org_id = 3328  and  addr.expire_date > sysdate and rad.std_addr_id = addr.std_addr_id) aa where aa.std_addr_name is not null";
            return OracleHelper.ExecuteNonQuery(sqlString);
        }
        #endregion

        #region 工程宽表数据
        public static bool gckb()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\工程\\FTTH光纤覆盖与入户信息表" + DateTime.Now.ToString("yyyyMMdd") + ".xls")) { return true; }
            else
            {
                string sqlString =
"select rad.jy_region_name 区域,\n" +
"       tmp.std_addr_name 小区,\n" +
"count(distinct case\n" +
"        when (subs.is_dtv = 1 and subs.is_paied = 1) or subs.is_lan_paied = 1 then\n" +
"         cust.cust_id\n" +
"        else\n" +
"         null\n" +
"      end) 缴费客户数,\n" +
            "       count(distinct case\n" +
"               when subs.is_dtv = 1 and subs.is_paied = 1 then\n" +
"                cust.cust_id\n" +
"               else\n" +
"                null\n" +
"             end) 数字缴费客户数,\n" +
"       count(distinct case\n" +
"               when subs.is_lan_paied = 1 then\n" +
"                cust.cust_id\n" +
"               else\n" +
"                null\n" +
"             end) 宽带缴费客户数,\n" +
"       count(distinct case\n" +
"               when exists\n" +
"                (select 1\n" +
"                       from cp2.cm_customer a\n" +
"                      where (a.remarks like 'ftth%' or a.remarks like 'FTTH%'  or a.remarks like 'Ftth%')\n" +
"                        and a.cust_id = cust.cust_id) then\n" +
"                cust.cust_id\n" +
"               else\n" +
"                null\n" +
"             end) 缴费客户数FTTH,\n" +
"       count(distinct case\n" +
"               when subs.is_dtv = 1 and subs.is_hdtv = 1 then\n" +
"                subs.bill_id\n" +
"               else\n" +
"                null\n" +
"             end) 高清机顶盒数\n" +
"  from cp2.cm_customer  cust,\n" +
"       rep2.rep_fact_um_subscriber_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddDays(-1).ToString("yyyyMMdd") + " subs,\n" +
"       wxjy.jy_region_address_rel           rad,\n" +
"       wxjy.jy_customer_ftth_rel            tmp\n" +
" where cust.cust_id = rad.cust_id\n" +
"   and cust.cust_id = tmp.cust_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and tmp.std_addr_name is not null\n" +
"   and cust.corp_org_id = 3328\n" +
" group by rad.jy_region_name, tmp.std_addr_name";

                DataTable dataTable = OracleHelper.ExecuteDataTable(sqlString);

                string sqlString1 =

    "select rad.jy_region_name 区域,\n" +
    "       tmp.std_addr_name 小区,\n" +
    "     count(distinct case\n" +
    "           when ofer.prod_service_id = 1002 and not exists\n" +
    "            (select 1\n" +
    "                   from files2.um_subscriber subs1,\n" +
    "                        files2.um_res        ures1,\n" +
    "                        res1.res_terminal    term1,\n" +
    "                        res1.res_sku         rsku1\n" +
    "                  where subs1.subscriber_ins_id = ures1.subscriber_ins_id\n" +
    "                    and ures1.res_equ_no = term1.serial_no\n" +
    "                    and term1.res_sku_id = rsku1.res_sku_id\n" +
    "                    and ures1.res_type_id = 2\n" +
    "                    and rsku1.res_sku_name in\n" +
    "                        ('银河高清基本型HDC6910(江阴)',\n" +
    "                         '银河高清交互型HDC691033(江阴)',\n" +
    "                         '银河智能高清交互型HDC6910798(江阴)',\n" +
    "                         '银河4K交互型HDC691090',\n" +
    "                         '4K超高清型II型融合型（EOC）',\n" +
    "                         '银河4K超高清II型融合型HDC6910B1(江阴)',\n" +
    "                         '4K超高清简易型（基本型）')\n" +
    "                    and subs1.cust_id = cust.cust_id) then\n" +
    "            cust.cust_id\n" +
    "           else\n" +
    "            null\n" +
    "         end) 纯标清客户缴费数,\n" +
    "       count(distinct case\n" +
    "               when ofer.prod_service_id = 1004 and exists\n" +
    "                (select 1\n" +
    "                       from upc1.pm_offer prod\n" +
    "                      where (prod.offer_name like '%300M%' or prod.offer_name like '%500M%' or  prod.offer_name like '%600M%'\n" +
    "                           or prod.offer_name like '%1000M%')\n" +
    "                        and prod.offer_id = fsta.offer_id) then\n" +
    "                cust.cust_id\n" +
    "               else\n" +
    "                null\n" +
    "             end) 缴费客户数FTTH按产品\n" +
    "  from cp2.cm_customer cust,\n" +
    "       files2.um_subscriber subs,\n" +
    "       files2.um_offer_06 ofer,\n" +
    "       files2.um_offer_sta_02 fsta,\n" +
    "       wxjy.jy_region_address_rel rad,\n" +
    "       wxjy.jy_customer_ftth_rel tmp\n" +
    " where cust.cust_id = rad.cust_id\n" +
    "  and  cust.cust_id = tmp.cust_id\n" +
    "   and cust.cust_id = subs.cust_id\n" +
    "   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
    "   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
    "   and fsta.expire_date > sysdate\n" +
    "   and ofer.expire_date > sysdate\n" +
    "   and ofer.prod_service_id in (1002, 1004)\n" +
    "   and fsta.offer_status = '1'\n" +
    "   and fsta.os_status is null\n" +
    "   and cust.own_corp_org_id = 3328\n" +
    "   and tmp.std_addr_name is not null\n" +
    " group by rad.jy_region_name, tmp.std_addr_name";


                DataTable dataTable1 = OracleHelper.ExecuteDataTable(sqlString1);

                DirectoryInfo pathInfo1 = new DirectoryInfo(Environment.CurrentDirectory);
                Excel.Application xapp = new Excel.Application();
                string filepath = pathInfo1.Parent.FullName + "\\工程源文件\\江阴有线电视网络覆盖用户信息表.xls";
                Excel.Workbook xbook = xapp.Workbooks._Open(filepath, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

                Excel.Worksheet xsheet = (Excel.Worksheet)xbook.Sheets["明细表"];

                DirectoryInfo pathInfo2 = new DirectoryInfo(Environment.CurrentDirectory);
                DataTable dt = ExcelHelper.InputFromExcel(pathInfo2.Parent.FullName + "\\工程源文件\\小区源.xls", "Sheet1");

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    string[] xq = dataTable.Rows[i]["小区"].ToString().Split(',');
                    if (dt.Rows.Count > 0)
                    {
                        for (int j = 0; j < dt.Rows.Count; j++)
                        {
                            if (dt.Rows[j]["ID"].ToString() == xq[0] && dt.Rows[j]["区域"].ToString() == dataTable.Rows[i]["区域"].ToString())//判断数据库重复结果中符合文档区域的数据
                            {
                                Excel.Range rng1 = xsheet.get_Range("K" + xq[0], Missing.Value);
                                rng1.Value2 = dataTable.Rows[i]["缴费客户数"].ToString();
                                Excel.Range rng2 = xsheet.get_Range("L" + xq[0], Missing.Value);
                                rng2.Value2 = dataTable.Rows[i]["数字缴费客户数"].ToString();
                                Excel.Range rng3 = xsheet.get_Range("M" + xq[0], Missing.Value);
                                rng3.Value2 = dataTable.Rows[i]["宽带缴费客户数"].ToString();
                                Excel.Range rng4 = xsheet.get_Range("N" + xq[0], Missing.Value);
                                rng4.Value2 = dataTable.Rows[i]["缴费客户数FTTH"].ToString();
                                Excel.Range rng10 = xsheet.get_Range("O" + xq[0], Missing.Value);
                                rng10.Value2 = dataTable.Rows[i]["高清机顶盒数"].ToString();
                            }
                        }
                    }
                }

                for (int i = 0; i < dataTable1.Rows.Count; i++)
                {
                    string[] xq = dataTable1.Rows[i]["小区"].ToString().Split(',');
                    if (dt.Rows.Count > 0)
                    {
                        for (int j = 0; j < dt.Rows.Count; j++)
                        {
                            if (dt.Rows[j]["ID"].ToString() == xq[0] && dt.Rows[j]["区域"].ToString() == dataTable1.Rows[i]["区域"].ToString())//判断数据库重复结果中符合文档区域的数据
                            {

                                Excel.Range rng11 = xsheet.get_Range("P" + xq[0], Missing.Value);
                                rng11.Value2 = dataTable1.Rows[i]["缴费客户数FTTH按产品"].ToString();
                                Excel.Range rng12 = xsheet.get_Range("Q" + xq[0], Missing.Value);
                                rng12.Value2 = dataTable1.Rows[i]["纯标清客户缴费数"].ToString();
                            }
                        }
                    }
                }


                Excel.Worksheet xsheet3 = (Excel.Worksheet)xbook.Sheets["Sheet1"];
                Excel.Range rng6 = xsheet3.get_Range("A1", Missing.Value);
                rng6.Value2 = sqlString;
                //Excel.Range rng6 = xsheet3.get_Range("A1", Missing.Value);
                //rng6.Value2 = sqlString.Substring(0, 8000);
                //Excel.Range rng7 = xsheet3.get_Range("B1", Missing.Value);
                //rng7.Value2 = sqlString.Substring(8000, 8000);
                //Excel.Range rng8 = xsheet3.get_Range("C1", Missing.Value);
                //rng8.Value2 = sqlString.Substring(16000, sqlString.Length - 16000);

                Excel.Worksheet xsheet8 = (Excel.Worksheet)xbook.Sheets["明细表"];
                Excel.Range rng9 = xsheet8.get_Range("A2", Missing.Value);
                rng9.Value2 = DateTime.Now.ToString("yyyy") + "年 " + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("MM") + "月 江阴有线电视网络FTTH光纤覆盖与入户明细表";

                // 2022年 4月 江阴有线电视网络FTTH覆盖用户明细


                string sqlString2 =

    "select rad.jy_region_name 区域,\n" +
    "       count(distinct case\n" +
    "               when subs.is_dtv = 1 and subs.is_paied = 1   then \n" +
    "                cust.cust_id\n" +
    "               else\n" +
    "                null\n" +
    "             end) 缴费客户数,\n" +
    "       count(distinct case\n" +
    "               when subs.is_lan_paied = 1 then\n" +
    "                cust.cust_id\n" +
    "               else\n" +
    "                null\n" +
    "             end) 宽带缴费客户数\n" +
    "  from cp2.cm_customer   cust,\n" +
    "       rep2.rep_fact_um_subscriber_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddDays(-1).ToString("yyyyMMdd") + " subs,\n" +
    "       wxjy.jy_region_address_rel           rad,\n" +
    "       wxjy.jy_region_sort_rel              rsr\n" +
    " where cust.cust_id = rad.cust_id\n" +
    "   and rsr.region_name = rad.jy_region_name\n" +
    "   and cust.cust_id = subs.cust_id\n" +
    "   and cust.corp_org_id = 3328\n" +
    " group by rad.jy_region_name";

                DataTable dataTable2 = OracleHelper.ExecuteDataTable(sqlString2);
                Excel.Worksheet xsheet1 = (Excel.Worksheet)xbook.Sheets["汇总表"];
                for (int i = 0; i < dataTable2.Rows.Count; i++)
                {
                    string str = dataTable2.Rows[i]["区域"].ToString();
                    if (str == "澄江东")
                    {
                        Excel.Range rng1 = xsheet1.get_Range("G6");
                        rng1.Value2 = dataTable2.Rows[i]["缴费客户数"].ToString();
                        Excel.Range rng2 = xsheet1.get_Range("L6");
                        rng2.Value2 = dataTable2.Rows[i]["宽带缴费客户数"].ToString();
                    }
                    if (str == "澄江")
                    {
                        Excel.Range rng3 = xsheet1.get_Range("G7");
                        rng3.Value2 = dataTable2.Rows[i]["缴费客户数"].ToString();
                        Excel.Range rng4 = xsheet1.get_Range("L7");
                        rng4.Value2 = dataTable2.Rows[i]["宽带缴费客户数"].ToString();
                    }
                    if (str == "澄江西")
                    {
                        Excel.Range rng5 = xsheet1.get_Range("G8");
                        rng5.Value2 = dataTable2.Rows[i]["缴费客户数"].ToString();
                        Excel.Range rng10 = xsheet1.get_Range("L8");
                        rng10.Value2 = dataTable2.Rows[i]["宽带缴费客户数"].ToString();
                    }
                    if (str == "高新区")
                    {
                        Excel.Range rng11 = xsheet1.get_Range("G9");
                        rng11.Value2 = dataTable2.Rows[i]["缴费客户数"].ToString();
                        Excel.Range rng12 = xsheet1.get_Range("L9");
                        rng12.Value2 = dataTable2.Rows[i]["宽带缴费客户数"].ToString();
                    }
                    if (str == "云亭")
                    {
                        Excel.Range rng13 = xsheet1.get_Range("G10");
                        rng13.Value2 = dataTable2.Rows[i]["缴费客户数"].ToString();
                        Excel.Range rng14 = xsheet1.get_Range("L10");
                        rng14.Value2 = dataTable2.Rows[i]["宽带缴费客户数"].ToString();
                    }
                    if (str == "周庄")
                    {
                        Excel.Range rng15 = xsheet1.get_Range("G11");
                        rng15.Value2 = dataTable2.Rows[i]["缴费客户数"].ToString();
                        Excel.Range rng16 = xsheet1.get_Range("L11");
                        rng16.Value2 = dataTable2.Rows[i]["宽带缴费客户数"].ToString();
                    }
                    if (str == "华士")
                    {
                        Excel.Range rng17 = xsheet1.get_Range("G12");
                        rng17.Value2 = dataTable2.Rows[i]["缴费客户数"].ToString();
                        Excel.Range rng18 = xsheet1.get_Range("L12");
                        rng18.Value2 = dataTable2.Rows[i]["宽带缴费客户数"].ToString();
                    }
                    if (str == "新桥")
                    {
                        Excel.Range rng19 = xsheet1.get_Range("G13");
                        rng19.Value2 = dataTable2.Rows[i]["缴费客户数"].ToString();
                        Excel.Range rng20 = xsheet1.get_Range("L13");
                        rng20.Value2 = dataTable2.Rows[i]["宽带缴费客户数"].ToString();
                    }
                    if (str == "顾山")
                    {
                        Excel.Range rng21 = xsheet1.get_Range("G14");
                        rng21.Value2 = dataTable2.Rows[i]["缴费客户数"].ToString();
                        Excel.Range rng22 = xsheet1.get_Range("L14");
                        rng22.Value2 = dataTable2.Rows[i]["宽带缴费客户数"].ToString();
                    }
                    if (str == "祝塘")
                    {
                        Excel.Range rng23 = xsheet1.get_Range("G15");
                        rng23.Value2 = dataTable2.Rows[i]["缴费客户数"].ToString();
                        Excel.Range rng24 = xsheet1.get_Range("L15");
                        rng24.Value2 = dataTable2.Rows[i]["宽带缴费客户数"].ToString();
                    }
                    if (str == "长泾")
                    {
                        Excel.Range rng25 = xsheet1.get_Range("G16");
                        rng25.Value2 = dataTable2.Rows[i]["缴费客户数"].ToString();
                        Excel.Range rng26 = xsheet1.get_Range("L16");
                        rng26.Value2 = dataTable2.Rows[i]["宽带缴费客户数"].ToString();
                    }
                    if (str == "南闸")
                    {
                        Excel.Range rng27 = xsheet1.get_Range("G17");
                        rng27.Value2 = dataTable2.Rows[i]["缴费客户数"].ToString();
                        Excel.Range rng28 = xsheet1.get_Range("L17");
                        rng28.Value2 = dataTable2.Rows[i]["宽带缴费客户数"].ToString();
                    }
                    if (str == "霞客")
                    {
                        Excel.Range rng29 = xsheet1.get_Range("G18");
                        rng29.Value2 = dataTable2.Rows[i]["缴费客户数"].ToString();
                        Excel.Range rng30 = xsheet1.get_Range("L18");
                        rng30.Value2 = dataTable2.Rows[i]["宽带缴费客户数"].ToString();
                    }
                    if (str == "月城")
                    {
                        Excel.Range rng30 = xsheet1.get_Range("G19");
                        rng30.Value2 = dataTable2.Rows[i]["缴费客户数"].ToString();
                        Excel.Range rng31 = xsheet1.get_Range("L19");
                        rng31.Value2 = dataTable2.Rows[i]["宽带缴费客户数"].ToString();
                    }
                    if (str == "青阳")
                    {
                        Excel.Range rng32 = xsheet1.get_Range("G20");
                        rng32.Value2 = dataTable2.Rows[i]["缴费客户数"].ToString();
                        Excel.Range rng33 = xsheet1.get_Range("L20");
                        rng33.Value2 = dataTable2.Rows[i]["宽带缴费客户数"].ToString();
                    }
                    if (str == "夏港")
                    {
                        Excel.Range rng34 = xsheet1.get_Range("G21");
                        rng34.Value2 = dataTable2.Rows[i]["缴费客户数"].ToString();
                        Excel.Range rng35 = xsheet1.get_Range("L21");
                        rng35.Value2 = dataTable2.Rows[i]["宽带缴费客户数"].ToString();
                    }
                    if (str == "申港")
                    {
                        Excel.Range rng36 = xsheet1.get_Range("G22");
                        rng36.Value2 = dataTable2.Rows[i]["缴费客户数"].ToString();
                        Excel.Range rng37 = xsheet1.get_Range("L22");
                        rng37.Value2 = dataTable2.Rows[i]["宽带缴费客户数"].ToString();
                    }
                    if (str == "利港")
                    {
                        Excel.Range rng38 = xsheet1.get_Range("G23");
                        rng38.Value2 = dataTable2.Rows[i]["缴费客户数"].ToString();
                        Excel.Range rng39 = xsheet1.get_Range("L23");
                        rng39.Value2 = dataTable2.Rows[i]["宽带缴费客户数"].ToString();
                    }
                    if (str == "璜土")
                    {
                        Excel.Range rng40 = xsheet1.get_Range("G24");
                        rng40.Value2 = dataTable2.Rows[i]["缴费客户数"].ToString();
                        Excel.Range rng41 = xsheet1.get_Range("L24");
                        rng41.Value2 = dataTable2.Rows[i]["宽带缴费客户数"].ToString();
                    }
                }
                Excel.Range rng42 = xsheet1.get_Range("A1");
                rng42.Value2 = DateTime.Now.ToString("yyyy") + "年 " + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("MM") + "月 江阴有线电视网络FTTH光纤覆盖与入户明细表";


                xbook.SaveAs(pathInfo.Parent.FullName + "\\工程\\FTTH光纤覆盖与入户信息表" + DateTime.Now.ToString("yyyyMMdd") + ".xls", Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                xsheet1 = null;
                xsheet = null;
                xbook = null;
                xapp.Quit();
                xapp = null;
                GC.Collect();

                return true;
            }
        }
        #endregion

        #region 工程FTTH缴费用户明细
        public static bool gc1()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\工程\\" + "工程FTTH缴费用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格,\n" +
"                tmp.std_addr_name 小区\n" +
"  from cp2.cm_customer                      cust,\n" +
"       cp2.cb_party                         part,\n" +
"       files2.um_address                    addr,\n" +
"       rep2.rep_fact_um_subscriber_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddDays(-1).ToString("yyyyMMdd") + " subs,\n" +
"       wxjy.jy_region_address_rel           rad,\n" +
"       wxjy.jy_contact_rel                  rel,\n" +
"       wxjy.jy_customer_ftth_rel            tmp,\n" +
"       wxjy.jy_customer_wg_rel              wg\n" +
" where cust.cust_id = rad.cust_id\n" +
"   and cust.cust_id = tmp.cust_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and tmp.std_addr_name is not null\n" +
"   and ((subs.is_dtv = 1 and subs.is_paied = 1) or subs.is_lan_paied = 1)\n" +
"   and (cust.remarks like 'ftth%' or cust.remarks like 'FTTH%' or cust.remarks like 'Ftth%')\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name, tmp.std_addr_name, cust.cust_code";

                DataTable dtt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 1 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\工程\\" + "工程FTTH缴费用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dtt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 工程全量用户明细
        public static bool gc2()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\工程\\" + "工程全量用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格,\n" +
"                tmp.std_addr_name 小区\n" +
"  from cp2.cm_customer                      cust,\n" +
"       cp2.cb_party                         part,\n" +
"       files2.um_address                    addr,\n" +
"       rep2.rep_fact_um_subscriber_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddDays(-1).ToString("yyyyMMdd") + " subs,\n" +
"       wxjy.jy_region_address_rel           rad,\n" +
"       wxjy.jy_contact_rel                  rel,\n" +
"       wxjy.jy_customer_ftth_rel            tmp,\n" +
"       wxjy.jy_customer_wg_rel              wg\n" +
" where cust.cust_id = rad.cust_id\n" +
"   and cust.cust_id = tmp.cust_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and tmp.std_addr_name is not null\n" +
"   and ((subs.is_dtv = 1 and subs.is_paied = 1) or subs.is_lan_paied = 1)\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name, tmp.std_addr_name, cust.cust_code";


                DataTable dtt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 1, 4 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\工程\\" + "工程全量用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dtt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 工程数字缴费用户明细
        public static bool gc3()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\工程\\" + "工程数字缴费用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格,\n" +
"                tmp.std_addr_name 小区\n" +
"  from cp2.cm_customer                      cust,\n" +
"       cp2.cb_party                         part,\n" +
"       files2.um_address                    addr,\n" +
"       rep2.rep_fact_um_subscriber_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddDays(-1).ToString("yyyyMMdd") + " subs,\n" +
"       wxjy.jy_region_address_rel           rad,\n" +
"       wxjy.jy_contact_rel                  rel,\n" +
"       wxjy.jy_customer_ftth_rel            tmp,\n" +
"       wxjy.jy_customer_wg_rel              wg\n" +
" where cust.cust_id = rad.cust_id\n" +
"   and cust.cust_id = tmp.cust_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and tmp.std_addr_name is not null\n" +
"   and (subs.is_dtv = 1 and subs.is_paied = 1)\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name, tmp.std_addr_name, cust.cust_code";
                DataTable dtt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 1, 4 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\工程\\" + "工程数字缴费用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dtt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 工程宽带缴费用户明细
        public static bool gc4()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\工程\\" + "工程宽带缴费用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格,\n" +
"                tmp.std_addr_name 小区\n" +
"  from cp2.cm_customer                      cust,\n" +
"       cp2.cb_party                         part,\n" +
"       files2.um_address                    addr,\n" +
"       rep2.rep_fact_um_subscriber_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddDays(-1).ToString("yyyyMMdd") + " subs,\n" +
"       wxjy.jy_region_address_rel           rad,\n" +
"       wxjy.jy_contact_rel                  rel,\n" +
"       wxjy.jy_customer_ftth_rel            tmp,\n" +
"       wxjy.jy_customer_wg_rel              wg\n" +
" where cust.cust_id = rad.cust_id\n" +
"   and cust.cust_id = tmp.cust_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and tmp.std_addr_name is not null\n" +
"   and subs.is_lan_paied = 1\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name, tmp.std_addr_name, cust.cust_code";


                DataTable dtt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 1, 4 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\工程\\" + "工程宽带缴费用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dtt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 工程纯标清客户缴费明细
        public static bool gc5()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\工程\\" + "工程纯标清客户缴费明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                term.serial_no 机顶盒号码,\n" +
"                subs.sub_bill_id 智能卡,\n" +
"                rsku.res_sku_name 资源类型,\n" +
"                case\n" +
"                  when nvl(subs.main_subscriber_ins_id, 0) = 0 then\n" +
"                   '主机'\n" +
"                  else\n" +
"                   ''\n" +
"                end 主副机,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格,\n" +
"                tmp.std_addr_name 小区,\n" +
"                cust.remarks 备注,\n" +
"                case\n" +
"                  when cust.remarks like 'ftth%' or\n" +
"                       cust.remarks like 'FTTH%' then\n" +
"                   '是'\n" +
"                  else\n" +
"                   ''\n" +
"                end 是否符合ftth小区\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_res              ures,\n" +
"       res1.res_terminal          term,\n" +
"       res1.res_sku               rsku,\n" +
"       files2.um_offer_06         ofer,\n" +
"       files2.um_offer_sta_02     fsta,\n" +
"       files2.um_address          addr,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_customer_wg_rel    wg,\n" +
"       wxjy.jy_customer_ftth_rel  tmp\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = tmp.cust_id\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"   and subs.subscriber_ins_id = ures.subscriber_ins_id\n" +
"   and ures.res_equ_no = term.serial_no\n" +
"   and term.res_sku_id = rsku.res_sku_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and nvl(subs.main_subscriber_ins_id, 0) = 0\n" +
"   and rsku.res_type_id = 2\n" +
"   and ofer.prod_service_id = 1002\n" +
"   and ures.expire_date > sysdate\n" +
"   and addr.expire_date > sysdate\n" +
"   and fsta.expire_date > sysdate\n" +
"   and ofer.expire_date > sysdate\n" +
"   and cust.own_corp_org_id = 3328\n" +
"   and not exists\n" +
" (select 1\n" +
"          from files2.um_subscriber subs1,\n" +
"               files2.um_res        ures1,\n" +
"               res1.res_terminal    term1,\n" +
"               res1.res_sku         rsku1\n" +
"         where subs1.subscriber_ins_id = ures1.subscriber_ins_id\n" +
"           and ures1.res_equ_no = term1.serial_no\n" +
"           and term1.res_sku_id = rsku1.res_sku_id\n" +
"           and ures1.res_type_id = 2\n" +
"           and rsku1.res_sku_name in\n" +
"               ('银河高清基本型HDC6910(江阴)',\n" +
"                '银河高清交互型HDC691033(江阴)',\n" +
"                '银河智能高清交互型HDC6910798(江阴)',\n" +
"                '银河4K交互型HDC691090',\n" +
"                '4K超高清型II型融合型（EOC）',\n" +
"                '银河4K超高清II型融合型HDC6910B1(江阴)',\n" +
"                '4K超高清简易型（基本型）',\n" +
"                '4K超高清IP机顶盒-便携款')\n" +
"           and subs1.cust_id = cust.cust_id)\n" +
"   and exists\n" +
" (select 1\n" +
"          from files2.um_offer_06 ofer2, files2.um_offer_sta_02 fsta2\n" +
"         where ofer2.offer_ins_id = fsta2.offer_ins_id\n" +
"           and fsta2.expire_date > sysdate\n" +
"           and ofer2.expire_date > sysdate\n" +
"           and fsta2.offer_status = '1'\n" +
"           and ofer2.prod_service_id = 1002\n" +
"           and fsta2.os_status is null\n" +
"           and ofer2.subscriber_ins_id = subs.subscriber_ins_id) --基本包为开通\n" +
" order by rad.jy_region_name,\n" +
"          addr.std_addr_name,\n" +
"          cust.cust_code,\n" +
"          term.serial_no";
                
                DataTable dtt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 1, 4, 5, 8 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\工程\\" + "工程纯标清客户缴费明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dtt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 工程高清机顶盒明细
        public static bool gc6()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\工程\\" + "工程高清机顶盒明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                term.serial_no 机顶盒,\n" +
"                subs.sub_bill_id 智能卡,\n" +
"                rsku.res_sku_name 类型,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格,\n" +
"                tmp.std_addr_name 小区\n" +
"  from cp2.cm_customer                      cust,\n" +
"       cp2.cb_party                         part,\n" +
"       files2.um_address                    addr,\n" +
"       rep2.rep_fact_um_subscriber_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddDays(-1).ToString("yyyyMMdd") + " subs,\n" +
"       wxjy.jy_region_address_rel           rad,\n" +
"       wxjy.jy_contact_rel                  rel,\n" +
"       wxjy.jy_customer_ftth_rel            tmp,\n" +
"       wxjy.jy_customer_wg_rel              wg,\n" +
"       res1.res_terminal                    term,\n" +
"       res1.res_sku                         rsku\n" +
" where cust.cust_id = rad.cust_id\n" +
"   and cust.cust_id = tmp.cust_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and subs.bill_id = term.serial_no\n" +
"   and term.res_sku_id = rsku.res_sku_id\n" +
"   and tmp.std_addr_name is not null\n" +
"   and subs.is_dtv = 1\n" +
"   and subs.is_hdtv = 1\n" +
"   and rsku.res_type_id = '2'\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name, tmp.std_addr_name, cust.cust_code";

                DataTable dtt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 1, 4, 5, 7 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\工程\\" + "工程高清机顶盒明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dtt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 工程FTTH按产品明细
        public static bool gc7()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\工程\\" + "工程FTTH按产品明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                tmp.std_addr_name 小区,\n" +
"                cust.remarks BOSS备注\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_address          addr,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_offer_06         ofer,\n" +
"       files2.um_offer_sta_02     fsta,\n" +
"       upc1.pm_offer              prod,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_customer_ftth_rel  tmp\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.cust_id = tmp.cust_id\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"   and fsta.offer_id = prod.offer_id\n" +
"   and subs.login_name is not null\n" +
"   and ofer.prod_service_id = 1004\n" +
"   and fsta.offer_status = '1'\n" +
"   and fsta.os_status is null\n" +
"   and addr.expire_date > sysdate\n" +
"   and fsta.expire_date > sysdate\n" +
"   and ofer.expire_date > sysdate\n" +
"   and prod.offer_name like '%FTTH%'\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.JY_REGION_NAME, addr.std_addr_name, cust.cust_code";
                DataTable dtt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 1, 4 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\工程\\" + "工程FTTH按产品明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dtt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion



        #region 日报订购客户级产品
        public static bool baobiao1()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "日报订购客户级产品--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")){ return true;}
            else
            {
                //   string oneday = "trunc(sysdate)-1";
                //if (DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).ToString("MMdd") == "1008")
                //{ oneday = "trunc(sysdate)-7"; }
                //else
                //{
                //    if (DateTime.Now.DayOfWeek.ToString() == "Monday")
                //    { oneday = "trunc(sysdate)-3"; }
                //}
                string sqlString =
    "select distinct cust.cust_code 客户证号,\n" +
    "                part.party_name 姓名,\n" +
    "                decode(cust.cust_type,\n" +
    "                       1,\n" +
    "                       '公众客户',\n" +
    "                       2,\n" +
    "                       '商业客户',\n" +
    "                       3,\n" +
    "                       '团体代付客户',\n" +
    "                       4,\n" +
    "                       '合同商业客户') 客户类型,\n" +
"                          case\n" +
"                  when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or \n" +
"                        (cust.cust_type = 4 and cust.cust_prop = 10)) then\n" +
"                   '业务用机(免催免停)'\n" +
"                  else\n" +
"                   '普通'\n" +
"                end 客户属性,\n" +
    "                '' 机顶盒,\n" +
    "                '' 智能卡,\n" +
    "                '' 资源型号,\n" +
    "                prod.offer_name 产品,\n" +
    "                '是' 客户级套餐产品,\n" +
    "                fsta.create_date 订购时间,\n" +
    "                fsta.valid_date 生效时间,\n" +
    "                staf.staff_name 操作员,\n" +
    "                orga.organize_name 营业员,\n" +
    "                rel.teleph_nunber 联系方式,\n" +
    "                rad.jy_region_name 区域,\n" +
    "                addr.std_addr_name 地址\n" +  
    "  from cp2.cm_customer            cust,\n" +
    "       cp2.cb_party               part,\n" +
    "       files2.um_subscriber       subs,\n" +
    "       files2.um_offer_06         ofer,\n" +
    "       files2.um_offer_sta_02     fsta,\n" +
    "       upc1.pm_offer              prod,\n" +
    "       wxjy.jy_region_address_rel rad,\n" +
    "       wxjy.jy_contact_rel        rel,\n" +
    "       files2.um_address          addr,\n" +
    "       params1.sec_operator       oper,\n" +
    "       params1.sec_staff          staf,\n" +
    "       params1.sec_organize       orga\n" +
    " where cust.cust_id = subs.cust_id\n" +
    "   and cust.cust_id = rad.cust_id(+)\n" +
    "   and cust.party_id = part.party_id\n" +
    "   and cust.cust_id = addr.cust_id(+)\n" +
    "   and cust.party_id = rel.party_id\n" +
    "   and cust.partition_id = rel.partition_id\n" +
    "   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
    "   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
    "   and fsta.offer_id = prod.offer_id\n" +
    "   and ofer.op_id = oper.operator_id(+)\n" +
    "   and oper.staff_id = staf.staff_id(+)\n" +
    "   and staf.organize_id = orga.organize_id(+)\n" +
    "   and subs.main_spec_id = 80020199\n" +
    "   and fsta.offer_status = '1'\n" +
    "   and fsta.os_status is null\n" +
    "   and (prod.offer_name like '%奇异%' or prod.offer_name like '%奇艺%' or\n" +
    "       prod.offer_name like '%芒果%' or prod.offer_name like '%学霸宝盒%' or\n" +
    "       prod.offer_name like '%电竞%' or prod.offer_name like '%致敬经典%' or\n" +
    "               prod.offer_name like '%探奇动物界%' or prod.offer_name like '%炫力%' or\n" +
    "               prod.offer_name like '%果果乐园%' or\n" +
    "       prod.offer_name like '%上文广%'  or  prod.offer_name like '%越剧%' or  prod.offer_name like '%中国互联网电视_8元/月_江苏_余额%')\n" +
    "   and fsta.expire_date > sysdate\n" +
    "   and ofer.expire_date > sysdate\n" +
    "   and addr.expire_date > sysdate\n" +
    "   and cust.own_corp_org_id = 3328";
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 1, 15 }; //哪些列是文本格式
                int[] columndate = { 10, 11, 12 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\日报\\" + "日报订购客户级产品--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 日报订购用户级产品
        public static bool baobiao2()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "日报订购用户级产品--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                //  string oneday = "trunc(sysdate)-1";
                //if (DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).ToString("MMdd") == "1008")
                //{ oneday = "trunc(sysdate)-7"; }
                //else
                //{
                //    if (DateTime.Now.DayOfWeek.ToString() == "Monday")
                //    { oneday = "trunc(sysdate)-3"; }
                //}
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                term.serial_no 机顶盒,\n" +
"                subs.sub_bill_id 智能卡,\n" +
"                rsku.res_sku_name 资源型号,\n" +
"                prod.offer_name 产品,\n" +
"                '' 客户级套餐产品,\n" +
"                fsta.create_date 订购时间,\n" +
"                fsta.valid_date 生效时间,\n" +
"                fsta.expire_date 失效时间,\n" +
"                staf.staff_name 操作员,\n" +
"                orga.organize_name 营业厅,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                          case\n" +
"                  when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or \n" +
"                        (cust.cust_type = 4 and cust.cust_prop = 10)) then\n" +
"                   '业务用机(免催免停)'\n" +
"                  else\n" +
"                   '普通'\n" +
"                end 客户属性,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址\n" + 
"  from cp2.cm_customer cust,\n" +
"       --files2.cm_account          acct,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_address          addr,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_res              ures,\n" +
"       res1.res_terminal          term,\n" +
"       res1.res_sku               rsku,\n" +
"       files2.um_offer_06         ofer,\n" +
"       files2.um_offer_sta_02     fsta,\n" +
"       upc1.pm_offer              prod,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       params1.sec_operator       oper,\n" +
"       params1.sec_staff          staf,\n" +
"       params1.sec_organize       orga\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ures.subscriber_ins_id\n" +
"   and ures.res_equ_no = term.serial_no\n" +
"   and term.res_sku_id = rsku.res_sku_id\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"   and fsta.offer_id = prod.offer_id\n" +
"   and ofer.op_id = oper.operator_id(+)\n" +
"   and oper.staff_id = staf.staff_id(+)\n" +
"   and staf.organize_id = orga.organize_id(+)\n" +
"   and rsku.res_type_id = 2\n" +
"   and fsta.offer_status = '1'\n" +
"   and fsta.os_status is null\n" +
"   and ures.expire_date > sysdate\n" +
"   and fsta.expire_date > sysdate\n" +
"   and （prod.offer_name like '%奇异%'  or prod.offer_name like '%奇艺%'  or prod.offer_name like '%炫力动漫%'   or prod.offer_name like '%芒果%'  or prod.offer_name like '%学霸宝盒%' \n" +
"    or prod.offer_name like '%电竞%'  or prod.offer_name like '%致敬经典%'  or prod.offer_name like '%上文广%' or prod.offer_name like '%极视%' or prod.offer_name like '%探奇动物界%' or prod.offer_name like '%果果乐园%' or prod.offer_name like '%电视机3年_0元/3年_江阴_首月出账%' or prod.offer_name like '%电视机5年_660元/3年_江阴_首月出账%'）\n" +
"   and cust.own_corp_org_id = 3328";
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 1, 4, 5, 14 }; //哪些列是文本格式
                int[] columndate = { 9, 10, 11 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\日报\\" + "日报订购用户级产品--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 日报退订产品
        public static bool baobiao3()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "日报退订产品--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                //    string oneday = "trunc(sysdate)-1";
                //if (DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).ToString("MMdd") == "1008")
                //{ oneday = "trunc(sysdate)-7"; }
                //else
                //{
                //    if (DateTime.Now.DayOfWeek.ToString() == "Monday")
                //    { oneday = "trunc(sysdate)-3"; }
                //}
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 客户姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                          case\n" +
"                  when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or \n" +
"                        (cust.cust_type = 4 and cust.cust_prop = 10)) then\n" +
"                   '业务用机(免催免停)'\n" +
"                  else\n" +
"                   '普通'\n" +
"                end 客户属性,\n" +
"                subs.bill_id 机顶盒,\n" +
"                subs.sub_bill_id 智能卡,\n" +
"                prod.offer_name 套餐,\n" +
"                ofer.offer_name 产品,\n" +
"                ofer.create_date 订购时间,\n" +
"                ofer.valid_date 生效时间,\n" +
"                ofer.expire_date 失效时间,\n" +
"                staf.staff_name 操作员,\n" +
"                orga.organize_name 营业厅,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址\n" +
"  from cp2.cm_customer cust,\n" +
"       files2.cm_account acct,\n" +
"       cp2.cb_party              part,\n" +
"(select *\n" +
"   from JOUR2.Om_Subscriber\n" +
" union all\n" +
" select *\n" +
"   from JOUR2.Om_Subscriber_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-2).ToString("yyyyMM") + "\n" +
" union all\n" +
" select *\n" +
"   from JOUR2.Om_Subscriber_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyyMM") + "\n" +
"    union all\n" +
" select *\n" +
"   from JOUR2.Om_Subscriber_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).ToString("yyyyMM") + ") subs,\n" +
"(select *\n" +
"   from JOUR2.OM_OFFER\n" +
"    union all\n" +
" select *\n" +
"   from JOUR2.OM_OFFER_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-2).ToString("yyyyMM") + "\n" +
"    union all\n" +
" select *\n" +
"   from JOUR2.OM_OFFER_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyyMM") + "\n" +
"    union all\n" +
" select *\n" +
"   from JOUR2.OM_OFFER_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).ToString("yyyyMM") + ") ofer,\n" +
"(select *\n" +
"   from jour2.om_order\n" +
" union all\n" +
" select *\n" +
"   from jour2.om_order_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-2).ToString("yyyyMM") + "\n" +
" union all\n" +
" select *\n" +
"   from jour2.om_order_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyyMM") + "\n" +
"    union all\n" +
" select *\n" +
"   from jour2.om_order_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).ToString("yyyyMM") + ") orde," +

"       params1.sec_operator oper,\n" +
"       params1.sec_staff staf,\n" +
"       params1.sec_organize orga,\n" +
"       wxjy.jy_contact_rel rel,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       files2.um_address addr,\n" +
"       upc1.pm_offer prod\n" +
" where cust.cust_id = acct.cust_id\n" +
"   and cust.cust_id = rad.cust_id\n" +
"   and cust.cust_id = addr.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = orde.party_role_id\n" +
"   and orde.order_id = subs.order_id\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.parent_offer_id = prod.offer_id\n" +
"   and ofer.op_id = oper.operator_id(+)\n" +
"   and oper.staff_id = staf.staff_id(+)\n" +
"   and staf.organize_id = orga.organize_id(+)\n" +
"   and ofer.offer_type <> 'S'\n" +
"   and ofer.action = 1\n" +
"   and ofer.expire_date < sysdate\n" +
//"      and  ofer.expire_date >= " + oneday + "\n" +
//"      and  ofer.expire_date < trunc(sysdate) \n" +
"   and (ofer.offer_name like '%奇异%' or ofer.offer_name like '%奇艺%' or ofer.offer_name like '%炫力动漫%' or ofer.offer_name like '%芒果TV%' or ofer.offer_name like '%学霸宝盒%' or ofer.offer_name like '%电竞%'\n" +
"    or ofer.offer_name like '%致敬经典%'  or ofer.offer_name like '%上文广%' or ofer.offer_name like '%探奇动物界%' or ofer.offer_name like '%果果乐园%' or ofer.offer_name like '%极视%'  or  ofer.offer_name like '%越剧%' or  ofer.offer_name like '%电视机3年_0元/3年_江阴_首月出账%' or  ofer.offer_name like '%电视机5年_660元/3年_江阴_首月出账%')\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by cust.cust_code, subs.bill_id, rad.jy_region_name";
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 1, 5, 6, 14 }; //哪些列是文本格式
                int[] columndate = { 9, 10, 11 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\日报\\" + "日报退订产品--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 日报云BOSS宽表综合统计
        public static bool baobiao4()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "日报云BOSS宽表综合统计--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct rad.jy_region_name 区域,\n" +
"count(distinct case when\n" +
"rfus.is_paied = 1 and rfus.is_dtv = 1  then rfus.cust_id END) 数字电视缴费用户1,\n" +
"count(distinct case when\n" +
"rfus.is_paied = 1 and rfus.is_dtv = 1 then rfus.subscriber_ins_id  END) 数字电视缴费终端数2,\n" +
"count(distinct case when\n" +
"rfus.is_dbitv=1 and rfus.is_dbitv_paied = 1 then  rfus.cust_id END) 互动缴费用户数3,\n" +
"count(distinct case when\n" +
"rfus.is_dbitv=1 and rfus.is_dbitv_paied = 1 then rfus.subscriber_ins_id END) 互动缴费终端数4,\n" +
"count(distinct case when\n" +
"rfus.is_hdtv=1  and rfus.is_dbitv_paied = 1 and rfus.is_dbitv=1  then  rfus.cust_id END) 高清互动缴费用户数5,\n" +
"count(distinct case when\n" +
"rfus.is_hdtv=1  and rfus.is_dbitv_paied = 1 and rfus.is_dbitv=1  then  rfus.subscriber_ins_id END) 高清互动缴费终端数6,\n" +
"count(distinct case when\n" +
"rfus.is_lan_paied = 1  then rfus.cust_id END) 宽带缴费用户数7,\n" +
"count(distinct case when\n" +
"rfus.is_lan_paied = 1  then rfus.subscriber_ins_id END) 宽带缴费终端数8,\n" +
"count(distinct case when\n" +
"rfus.is_valid1 = 1 and rfus.is_dtv = 1 then rfus.cust_id END) 数字电视有效用户数9,\n" +
"count(distinct case when\n" +
"rfus.is_valid1 = 1 and rfus.is_dtv = 1 then rfus.subscriber_ins_id END) 数字电视有效终端数10,-- 李诚：is_valid1好像是欠费停机1年以内的，2是两年\n" +
"count(distinct case when cust.cust_type=1   --住宅\n" +
"and rfus.is_valid1 = 1 and rfus.is_dtv = 1 then rfus.cust_id END) 住宅数字电视有效用户数11,   --11的单独跑好了\n" +
"count(distinct case when\n" +
"rfus.is_dbitv=1 and rfus.is_valid1 = 1 then rfus.cust_id END) 互动有效用户数12,\n" +
"count(distinct case when\n" +
"rfus.is_dbitv=1 and rfus.is_valid1 = 1 then rfus.subscriber_ins_id END) 互动有效终端数13,\n" +
"count(distinct case when\n" +
"rfus.is_dbitv=1 and rfus.is_valid1= 1 and rfus.is_hdtv = 1   then rfus.cust_id END) 高清互动有效用户数14,\n" +
"count(distinct case when\n" +
"rfus.is_dbitv=1 and rfus.is_valid1= 1 and rfus.is_hdtv = 1 then rfus.subscriber_ins_id END) 高清互动有效终端数15,\n" +
"count(distinct case when\n" +
"rfus.is_lan_valid = 1 then rfus.cust_id END) 宽带有效用户数16,\n" +
"count(distinct case when\n" +
"rfus.is_lan_valid = 1 then rfus.subscriber_ins_id END) 宽带有效终端数17,\n" +
"count(distinct case when  cust.cust_type <> 1   --非住宅\n" +
"and rfus.is_paied = 1 and rfus.is_dtv = 1\n" +
"and rfus.user_name not like '%test%'\n" +
"then rfus.subscriber_ins_id  END) 非住宅数字电视缴费终端数18,\n" +
"count(DISTINCT case when rfus.is_4k=1  and rfus.is_dbitv=1 and rfus.is_dbitv_paied = 1  then rfus.subscriber_ins_id end)  高清4k互动缴费终端19,\n" +
"count(DISTINCT case when rfus.is_4k=1  and rfus.is_dbitv=1 and rfus.is_valid1= 1  then rfus.subscriber_ins_id end)   高清4k互动有效终端20,\n" +
"count(DISTINCT case when rfus.is_4k=1  and rfus.is_paied = 1 and rfus.is_dtv=1 then rfus.subscriber_ins_id end)  高清4k数字电视缴费终端21,\n" +
"count(DISTINCT case when rfus.is_4k=1  and rfus.is_dtv=1 and rfus.is_valid1= 1 then rfus.subscriber_ins_id end)  高清4k数字电视有效终端22,\n" +
"count(DISTINCT case when rfus.is_dtv=1 and rfus.is_4k = 1 and rfus.is_paied = 1 then rfus.cust_id  end)  高清4K缴费用户数23,\n" +
"count(DISTINCT case when rfus.is_dtv=1 and rfus.is_4k = 1 and rfus.is_valid1= 1  then rfus.cust_id  end)  高清4k有效用户数24\n" +
"/*count(DISTINCT case when rfus.is_hdtv=1 and rfus.is_4k = 0 and rfus.is_paied = 1 then rfus.cust_id  end)  高清非4K缴费用户数25,\n" +
"count(DISTINCT case when rfus.is_hdtv=1 and rfus.is_4k = 0 and rfus.is_valid1= 1  then rfus.cust_id  end)  高清非4k有效用户数26*/\n" +
"from  rep2.rep_fact_um_subscriber_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddDays(-1).ToString("yyyyMMdd") + "  rfus, cp2.cm_customer cust, wxjy.jy_region_address_rel rad\n" +
"where cust.cust_id = rfus.cust_id and cust.cust_id = rad.cust_id\n" +
"and rfus.corp_org_id = 3328\n" +
"group by rad.jy_region_name\n" +
"order by rad.jy_region_name";
                int[] columntxt = { 0 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\日报\\" + "日报云BOSS宽表综合统计--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 日报订购芒果等数字包年产品明细
        public static bool baobiao5()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "日报订购芒果等数字包年产品明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                          case\n" +
"                  when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or \n" +
"                        (cust.cust_type = 4 and cust.cust_prop = 10)) then\n" +
"                   '业务用机(免催免停)'\n" +
"                  else\n" +
"                   '普通'\n" +
"                end 客户属性,\n" +
"                term.serial_no 机顶盒,\n" +
"                subs.sub_bill_id 智能卡,\n" +
"                rsku.res_sku_name 资源型号,\n" +
"                prod1.offer_name 套餐,\n" +
"                prod.offer_name 产品,\n" +
"                fsta.create_date 订购时间,\n" +
"                fsta.valid_date 生效时间,\n" +
"                fsta.expire_date 失效时间,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_address          addr,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_res              ures,\n" +
"       res1.res_terminal          term,\n" +
"       res1.res_sku               rsku,\n" +
"       files2.um_offer_06         ofer,\n" +
"       files2.um_offer_sta_02     fsta,\n" +
"       upc1.pm_offer              prod,\n" +
"       upc1.pm_offer              prod1,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_customer_wg_rel    wg\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ures.subscriber_ins_id\n" +
"   and ures.res_equ_no = term.serial_no\n" +
"   and term.res_sku_id = rsku.res_sku_id\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"   and fsta.offer_id = prod.offer_id\n" +
"   and fsta.parent_offer_id = prod1.offer_id\n" +
"   and rsku.res_type_id = 2\n" +
"   and subs.state = '1'\n" +
"   and fsta.offer_status = '1'\n" +
"   and fsta.os_status is null\n" +
"   and addr.expire_date > sysdate\n" +
"   and ures.expire_date > sysdate\n" +
"   and fsta.expire_date > sysdate\n" +
"   and ofer.expire_date > sysdate\n" +
"   and prod.offer_name in ('芒果TV首年288元_江阴',\n" +
"                           '4K套餐B首年480元_江阴',\n" +
"                           '极视影院_4K套餐包年300元_江阴',\n" +
"                           '4K套餐C首年480元_江阴',\n" +
"                           '4K套餐E首年360元_江阴')\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name,\n" +
"          addr.std_addr_name,\n" +
"          cust.cust_code,\n" +
"          term.serial_no";
                int[] columntxt = { 1, 5, 6, 13 }; //哪些列是文本格式
                int[] columndate = { 10, 11, 12 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\日报\\" + "日报订购芒果等数字包年产品明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 日报订购FTTH宽带升级购机套餐明细
        public static bool baobiao6()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "日报订购FTTH宽带升级购机套餐明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                subs.login_name 宽带登录名,\n" +
"                prod1.offer_name 套餐,\n" +
"                prod.offer_name 产品,\n" +
"                fsta.create_date 订购时间,\n" +
"                fsta.valid_date 生效时间,\n" +
"                fsta.expire_date 失效时间,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格,\n" +
"                          case\n" +
"                  when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or \n" +
"                        (cust.cust_type = 4 and cust.cust_prop = 10)) then\n" +
"                   '业务用机(免催免停)'\n" +
"                  else\n" +
"                   '普通'\n" +
"                end 客户属性\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_address          addr,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_offer_06         ofer,\n" +
"       files2.um_offer_sta_02     fsta,\n" +
"       upc1.pm_offer              prod,\n" +
"       upc1.pm_offer              prod1,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_customer_wg_rel    wg\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"   and fsta.offer_id = prod.offer_id\n" +
"   and fsta.parent_offer_id = prod1.offer_id\n" +
"   and subs.state = '1'\n" +
"   and fsta.offer_status = '1'\n" +
"   and fsta.os_status is null\n" +
"   and addr.expire_date > sysdate\n" +
"   and fsta.expire_date > sysdate\n" +
"   and ofer.expire_date > sysdate\n" +
"   and prod.offer_name = 'FTTH宽带升级购机套餐_江阴(首月96元)'\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name,\n" +
"          addr.std_addr_name,\n" +
"          cust.cust_code,\n" +
"          subs.login_name";

                int[] columntxt = { 1, 10 }; //哪些列是文本格式
                int[] columndate = { 7, 8, 9 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\日报\\" + "日报订购FTTH宽带升级购机套餐明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 日报订购客户级时间量产品
        public static bool baobiao7()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "日报订购客户级时间量产品--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct cust.cust_code 客户证号,\n" +
"                        part.party_name 客户姓名,\n" +
"                        decode(cust.cust_type,\n" +
"                               1,\n" +
"                               '公众客户',\n" +
"                               2,\n" +
"                               '商业客户',\n" +
"                               3,\n" +
"                               '团体代付客户',\n" +
"                               4,\n" +
"                               '合同商业客户') 客户类型,\n" +
"                          case\n" +
"                  when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or\n" +
"                        (cust.cust_type = 4 and cust.cust_prop = 10)) then\n" +
"                   '业务用机(免催免停)'\n" +
"                  else\n" +
"                   '普通'\n" +
"                end 客户属性,\n" +
"                        '' 机顶盒,\n" +
"                        '' 智能卡,\n" +
"                        '' 资源型号,\n" +
"                        ofer.offer_name 产品,\n" +
"                        '是' 客户级套餐产品,\n" +
"                        ofer.create_date 订购时间,\n" +
"                        ofer.valid_date 生效时间,\n" +
"                        ofer.expire_date 失效时间,\n" +
"                        staf.staff_name 操作员,\n" +
"                        orga.organize_name 营业员,\n" +
"                        rel.teleph_nunber 联系方式,\n" +
"                        rad.jy_region_name 区域,\n" +
"                        addr.std_addr_name 地址\n" +
"          from cp2.cm_customer            cust,\n" +
"               files2.um_subscriber       subs,\n" +
"               files2.um_offer_06         ofer,\n" +
"               wxjy.jy_region_address_rel rad,\n" +
"               wxjy.jy_contact_rel        rel,\n" +
"               files2.um_address          addr,\n" +
"               params1.sec_operator       oper,\n" +
"               params1.sec_staff          staf,\n" +
"       cp2.cb_party              part,\n" +
"               params1.sec_organize       orga\n" +
"         where cust.cust_id = subs.cust_id\n" +
"           and cust.cust_id = rad.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"           and cust.cust_id = addr.cust_id\n" +
"           and cust.party_id = rel.party_id\n" +
"           and cust.partition_id = rel.partition_id\n" +
"           and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"           and ofer.op_id = oper.operator_id(+)\n" +
"           and oper.staff_id = staf.staff_id(+)\n" +
"           and staf.organize_id = orga.organize_id(+)\n" +
"           and ofer.parent_offer_id <> -1            and addr.expire_date>sysdate\n" +
"           and subs.main_spec_id = 80020199\n" +
"           and (ofer.offer_name like '%炫力%' or ofer.offer_name like '%极视%' or ofer.offer_name like '%芒果%' or ofer.offer_name like '%动物%'  or ofer.offer_name like '%亲子%' or ofer.offer_name like '%奇异%')\n" +
"           and ofer.expire_date > sysdate\n" +
"           and cust.own_corp_org_id = 3328";
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 1, 15 }; //哪些列是文本格式
                int[] columndate = { 10, 11, 12 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\日报\\" + "日报订购客户级时间量产品--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 日报通用现金账本余额不足20元的客户明细
        public static bool baobiao8()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "日报通用现金账本余额不足20元的客户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct   cust.cust_code 客户证号,\n" +
"                tmp.acct_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                tmp.asset_item_name 账本,\n" +
"                tmp.amount 余额,\n" +
"                case\n" +
"                  when exists (select 1\n" +
"                          from files2.um_subscriber subs1\n" +
"                         where subs1.main_spec_id = 80020003\n" +
"                           and subs1.cust_id = cust.cust_id) then\n" +
"                   '是'\n" +
"                  else\n" +
"                   ''\n" +
"                end 是否有宽带,\n" +
"                case\n" +
"                  when exists (select 1\n" +
"                          from files2.um_subscriber subs2,\n" +
"                               files2.um_res        ures2,\n" +
"                               res2.res_terminal    term2,\n" +
"                               res2.res_sku         rsku2\n" +
"                         where subs2.subscriber_ins_id =\n" +
"                               ures2.subscriber_ins_id\n" +
"                           and ures2.res_equ_no = term2.serial_no\n" +
"                           and term2.res_sku_id = rsku2.res_sku_id\n" +
"                           and rsku2.res_type_id = 2\n" +
"                           and ures2.expire_date > sysdate\n" +
"                           and rsku2.res_sku_name in\n" +
"                               ('银河高清基本型HDC6910(江阴)',\n" +
"                                '银河高清交互型HDC691033(江阴)',\n" +
"                                '银河智能高清交互型HDC6910798(江阴)',\n" +
"                                '银河4K交互型HDC691090',\n" +
"                                '4K超高清型II型融合型（EOC）',\n" +
"                                '银河4K超高清II型融合型HDC6910B1(江阴)',\n" +
"                                '4K超高清简易型（基本型）')\n" +
"                           and subs2.cust_id = cust.cust_id) then\n" +
"                   '有'\n" +
"                  else\n" +
"                   ''\n" +
"                end 是否存在高清机顶盒,\n" +
"                 wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格\n" +
"  from cp2.cm_customer cust,\n" +
"       (select acct.cust_id,\n" +
"               acct.acct_id,\n" +
"               acct.acct_name,\n" +
"               nvl(aaty.asset_item_id, 100) asset_item_id,\n" +
"               nvl(aaty.asset_item_name, '通用现金账本') asset_item_name,\n" +
"               nvl(sum(blac.balance), 0) / 100 amount\n" +
"          from files2.cm_account  acct,\n" +
"               ac2.am_balance_" + DateTime.Now.ToString("MM") + "  blac, --每月一张表\n" +
"               pzg1.am_asset_type aaty\n" +
"         where acct.acct_id = blac.acct_id(+)\n" +
"           and blac.asset_item_id = aaty.asset_item_id(+)\n" +
"           and acct.acct_name not like '%测试%'\n" +
"           and acct.acct_name not like '%ceshi%'\n" +
"           and acct.acct_name not like '%test%'\n" +
"           and acct.corp_org_id = 3328\n" +
"         group by acct.cust_id,\n" +
"                  acct.acct_id,\n" +
"                  acct.acct_name,\n" +
"                  nvl(aaty.asset_item_id, 100),\n" +
"                  nvl(aaty.asset_item_name, '通用现金账本')) tmp,\n" +
"       wxjy.jy_region_address_rel rad, wxjy.jy_customer_wg_rel    wg,\n" +
"       files2.um_address addr,\n" +
"       wxjy.jy_contact_rel rel\n" +
" where cust.cust_id = tmp.cust_id\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = rad.cust_id\n" +
"   and cust.cust_id = addr.cust_id\n" +
"   and cust.cust_type = 1 --只统计公众客户\n" +
"   and cust.cust_prop <> 6 --去除免费客户\n" +
"   and cust.cust_prop <> 13 --去除广联客户\n" +
"   and tmp.asset_item_id = 100 --通用现金账本\n" +
"   and addr.expire_date > sysdate\n" +
"   and tmp.amount >= 0\n" +
"   and tmp.amount < 20\n" +
"   and cust.own_corp_org_id = 3328   and cust.cust_id=wg.cust_id\n" +
"   and not exists (select 1\n" +
"          from ac2.am_entrust_relation rela\n" +
"         where rela.exp_date > sysdate\n" +
"           and rela.state = 0\n" +
"           and rela.acct_id = tmp.acct_id) --不是银行托收客户\n" +
"   and exists (select 1\n" +
"          from files2.um_subscriber   subs,\n" +
"               files2.um_offer_06     ofer,\n" +
"               files2.um_offer_sta_02 fsta\n" +
"         where subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"           and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"           and subs.main_spec_id = 80020001\n" +
"           and ofer.expire_date > sysdate\n" +
"           and ofer.prod_service_id = 1002\n" +
"           and fsta.offer_status = '1'\n" +
"           and fsta.os_status is null\n" +
"           and subs.cust_id = cust.cust_id) --有开通的机顶盒\n" +
" order by rad.jy_region_name, addr.std_addr_name, cust.cust_code";

                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 1, 4, 6 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\日报\\" + "日报通用现金账本余额不足20元的客户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 日报上个月欠停且当前未复通的用户明细
        public static bool baobiao9()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "日报上个月欠停且当前未复通的用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                fsta.done_date 受理时间,\n" +
"                sum(qf.balance) / 100 历史账单欠费金额,\n" +
"                decode(rela.state, 0, '是', 1, '') 是否银行代扣,\n" +
"                case\n" +
"                  when exists (select 1\n" +
"                          from files2.um_subscriber subs2,\n" +
"                               files2.um_res        ures2,\n" +
"                               res1.res_terminal    term2,\n" +
"                               res1.res_sku         rsku2\n" +
"                         where subs2.subscriber_ins_id =\n" +
"                               ures2.subscriber_ins_id\n" +
"                           and ures2.res_equ_no = term2.serial_no\n" +
"                           and term2.res_sku_id = rsku2.res_sku_id\n" +
"                           and rsku2.res_type_id = 2\n" +
"                           and ures2.expire_date > sysdate\n" +
"                           and rsku2.res_sku_name in\n" +
"                               ('银河高清基本型HDC6910(江阴)',\n" +
"                                '银河高清交互型HDC691033(江阴)',\n" +
"                                '银河智能高清交互型HDC6910798(江阴)',\n" +
"                                '银河4K交互型HDC691090',\n" +
"                                '4K超高清型II型融合型（EOC）',\n" +
"                                '银河4K超高清II型融合型HDC6910B1(江阴)',\n" +
"                                '4K超高清简易型（基本型）')\n" +
"                           and subs2.cust_id = cust.cust_id) then\n" +
"                   '有'\n" +
"                  else\n" +
"                   ''\n" +
"                end 是否存在高清机顶盒,\n" +
"                rela.bank_name 开户行,\n" +
"                rela.bank_acct_id 开户账号,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.cm_account          acct,\n" +
"       ac2.am_bill_item           qf,\n" +
"       files2.um_address          addr,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_customer_wg_rel    wg,\n" +
"       ac2.am_entrust_relation    rela,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_offer_06         ofer,\n" +
"       files2.um_offer_sta_02     fsta\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = acct.cust_id\n" +
"   and acct.acct_id = qf.acct_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and acct.acct_id = rela.acct_id(+)\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"   and nvl(subs.main_subscriber_ins_id, 0) = 0\n" +
"   and ofer.prod_service_id = 1002\n" +
"   and fsta.os_status = '1'\n" +
"   and cust.cust_prop <> 6\n" +
"   and part.party_name not like '%测试%'\n" +
"   and part.party_name not like '%ceshi%'\n" +
"   and part.party_name not like '%test%'\n" +
"   and addr.expire_date > sysdate\n" +
"   and ofer.expire_date > sysdate\n" +
"   and fsta.expire_date > sysdate\n" +
"   and cust.own_corp_org_id = 3328\n" +
"   and  fsta.done_date>=date'" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"   and  fsta.done_date<date'" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"   and not exists\n" +
" (select 1\n" +
"          from files2.um_offer_06 ofer1, files2.um_offer_sta_02 fsta1\n" +
"         where ofer1.offer_ins_id = fsta1.offer_ins_id\n" +
"           and ofer1.prod_service_id = 1002\n" +
"           and fsta1.offer_status = '1'\n" +
"           and fsta1.os_status is null\n" +
"           and fsta1.expire_date > sysdate\n" +
"           and ofer1.expire_date > sysdate\n" +
"           and ofer1.subscriber_ins_id = subs.subscriber_ins_id)\n" +
" group by cust.cust_code,cust.cust_id,\n" +
"          part.party_name,\n" +
"          cust.cust_type,\n" +
"          rela.state,\n" +
"          rela.bank_name,\n" +
"          rela.bank_acct_id,\n" +
"          rel.teleph_nunber,\n" +
"          rad.jy_region_name,\n" +
"          addr.std_addr_name,\n" +
"          wg.region_name,\n" +
"          wg.grid_name,\n" +
"          fsta.done_date\n" +
" order by rad.jy_region_name, addr.std_addr_name, cust.cust_code";

                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 1, 9, 10 }; //哪些列是文本格式
                int[] columndate = { 4 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\日报\\" + "日报上个月欠停且当前未复通的用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion
        
        #region 日报1月1号至今的缴费客户明细数据
        public static bool baobiao11()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "日报1月1号至今的缴费客户明细数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                cust.create_date 开户时间,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_address          addr,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_customer_wg_rel    wg\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and addr.expire_date > sysdate\n" +
"   and cust.create_date >= date '2023-1-1'\n" +
"    and cust.create_date < date '" + DateTime.Now.ToString("yyyy-MM-dd") + "'\n" +
"   and cust.own_corp_org_id = 3328\n" +
"   and exists (select 1\n" +
"          from files2.um_subscriber   subs,\n" +
"               files2.um_offer_06     ofer,\n" +
"               files2.um_offer_sta_02 fsta\n" +
"         where subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"           and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"           and ofer.prod_service_id = 1002\n" +
"           and fsta.offer_status = '1'\n" +
"           and fsta.os_status is null\n" +
"           and fsta.expire_date > sysdate\n" +
"           and ofer.expire_date > sysdate\n" +
"           and subs.cust_id = cust.cust_id)\n" +
"   and not exists (select 1\n" +
"          from rep2.rep_fact_cm_customer_20221231 a\n" +
"         where a.cust_id = cust.cust_id)\n" +
" order by rad.jy_region_name, addr.std_addr_name, cust.cust_code";


                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 1, 5 }; //哪些列是文本格式
                int[] columndate = { 4 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\日报\\" + "日报1月1号至今的缴费客户明细数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion
        
        #region 日报宽带用户级产品订购
        public static bool baobiao13()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "日报宽带用户级产品订购--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 客户姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                subs.login_name 宽带登录名,\n" +
"                updf.offer_name 产品名称,\n" +
"                fsta.valid_date 生效时间,\n" +
"                 wgrel.grid_name 网格,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                          case\n" +
"                  when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or \n" +
"                        (cust.cust_type = 4 and cust.cust_prop = 10)) then\n" +
"                   '业务用机(免催免停)'\n" +
"                  else\n" +
"                   '普通'\n" +
"                end 客户属性\n" +
"  from cp2.cm_customer            cust,\n" +
"       files2.cm_account          acct,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_offer_06         ofer,\n" +
"       files2.um_offer_sta_02     fsta,\n" +
"       upc1.pm_offer              updf,\n" +
"       wxjy.jy_customer_wg_rel              wgrel,\n" +
"        wxjy.jy_region_address_rel rad,\n" +
"       cp2.cb_party              part,\n" +
"         files2.um_address          addr\n" +
" where cust.cust_id = subs.cust_id\n" +
"   and cust.cust_id = acct.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"    and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"  and cust.cust_id=wgrel.cust_id(+)\n" +
"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"   and ofer.subscriber_ins_id = subs.subscriber_ins_id\n" +
"   and fsta.offer_id = updf.offer_id\n" +
"   and subs.state = '1'\n" +

"     and addr.expire_date > sysdate\n" +
"   and  fsta.expire_date > sysdate\n" +
" and (updf.offer_name like '%100M%'  or updf.offer_name like '%200M%'  or updf.offer_name like '%300M%'  or updf.offer_name like '%500M%'  or updf.offer_name like '%600M%'  or  updf.offer_name like '%1000M%' or  updf.offer_name like '%电视机%' or  updf.offer_name like '%广电监控服务%' or  updf.offer_name like '%宽带体验升级%')\n" +
"   and fsta.os_status is null\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by cust.cust_code";

                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 1, 4 }; //哪些列是文本格式
                int[] columndate = { 6 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\日报\\" + "日报宽带用户级产品订购--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 日报宽带产品退订
        public static bool baobiao14()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "日报宽带产品退订--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 客户姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                          case\n" +
"                  when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or \n" +
"                        (cust.cust_type = 4 and cust.cust_prop = 10)) then\n" +
"                   '业务用机(免催免停)'\n" +
"                  else\n" +
"                   '普通'\n" +
"                end 客户属性,\n" +
"                prod.offer_name 套餐,\n" +
"                ofer.offer_name 产品,\n" +
"                ofer.create_date 订购时间,\n" +
"                ofer.valid_date 生效时间,\n" +
"                ofer.expire_date 失效时间,\n" +
"                staf.staff_name 操作员,\n" +
"                orga.organize_name 营业厅,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址\n" +
"  from cp2.cm_customer cust,\n" +
"       files2.cm_account acct,\n" +
"       cp2.cb_party              part,\n" +
"(select *\n" +
"   from JOUR2.Om_Subscriber\n" +
" union all\n" +
" select *\n" +
"   from JOUR2.Om_Subscriber_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "\n" +
"    union all\n" +
" select *\n" +
"   from JOUR2.Om_Subscriber_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).ToString("yyyyMM") + ") subs,\n" +
"(select *\n" +
"   from JOUR2.OM_OFFER\n" +
"    union all\n" +
" select *\n" +
"   from JOUR2.OM_OFFER_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "\n" +
"    union all\n" +
" select *\n" +
"   from JOUR2.OM_OFFER_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).ToString("yyyyMM") + ") ofer,\n" +
"(select *\n" +
"   from jour2.om_order\n" +
" union all\n" +
" select *\n" +
"   from jour2.om_order_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "\n" +
"    union all\n" +
" select *\n" +
"   from jour2.om_order_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).ToString("yyyyMM") + ") orde," +
"       params1.sec_operator oper,\n" +
"       params1.sec_staff staf,\n" +
"       params1.sec_organize orga,\n" +
"       wxjy.jy_contact_rel rel,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       files2.um_address addr,\n" +
"       upc1.pm_offer prod\n" +
" where cust.cust_id = acct.cust_id\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = orde.party_role_id\n" +
"   and orde.order_id = subs.order_id\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.parent_offer_id = prod.offer_id\n" +
"   and ofer.op_id = oper.operator_id(+)\n" +
"   and oper.staff_id = staf.staff_id(+)\n" +
"   and staf.organize_id = orga.organize_id(+)\n" +
"   and ofer.offer_type <> 'S'\n" +
"   and ofer.action = 1\n" +
"   and (ofer.offer_name like '%100M%' or ofer.offer_name like '%200M%' or ofer.offer_name like '%300M%' or ofer.offer_name like '%500M%'  or ofer.offer_name like '%600M%' or ofer.offer_name like '%1000M%'  or  ofer.offer_name like '%电视机%'  or  ofer.offer_name like '%广电监控服务10元/月%'  or  ofer.offer_name like '%广电监控服务15元/月%'  or  ofer.offer_name like '%广电监控服务首月180元-次年10元/月%')\n" +
"   and cust.own_corp_org_id = 3328";

                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 1, 12 }; //哪些列是文本格式
                int[] columndate = { 7, 8, 9 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\日报\\" + "日报宽带产品退订--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 日报宽带产品变更的终端明细
        public static bool baobiao15()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "日报宽带产品变更的终端明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                subs.login_name 宽带登录名,\n" +
"                prod1.offer_name 变更前的宽带产品,\n" +
"                prod2.offer_name 变更后的宽带产品,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格,\n" +
"                          case\n" +
"                  when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or \n" +
"                        (cust.cust_type = 4 and cust.cust_prop = 10)) then\n" +
"                   '业务用机(免催免停)'\n" +
"                  else\n" +
"                   '普通'\n" +
"                end 客户属性\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_address          addr,\n" +
"       files2.um_subscriber       subs,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_customer_wg_rel    wg,\n" +
"       rep2.rep_fact_ins_srvpkg_20230101 insp1,\n" +
"       upc1.pm_offer                     prod1,\n" +
"       rep2.rep_fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddDays(-1).ToString("yyyyMMdd") + " insp2,\n" +
"       upc1.pm_offer                     prod2\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.party_id = rel.party_id(+)\n" +
"   and cust.partition_id = rel.partition_id(+)\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = insp1.subscriber_ins_id\n" +
"   and insp1.srvpkg_id = prod1.offer_id\n" +
"   and insp1.prod_service_id = 1004\n" +
"   and subs.subscriber_ins_id = insp2.subscriber_ins_id\n" +
"   and insp2.srvpkg_id = prod2.offer_id\n" +
"   and insp2.prod_service_id = 1004\n" +
"   and prod1.offer_id <> prod2.offer_id\n" +
"   and subs.login_name is not null\n" +
"   and subs.state <> 'D'\n" +
"   and addr.expire_date > sysdate\n" +
"   and cust.corp_org_id = 3328\n" +
" order by addr.std_addr_name, cust.cust_code, subs.login_name";
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 1, 4, 7 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\日报\\" + "日报宽带产品变更的终端明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 日报云BOSS宽表综合统计带网格人员
        public static bool baobiao16()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "日报云BOSS宽表综合统计带网格人员--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct rad.jy_region_name 区域, wg.grid_name 网格, wg.mgr_name 网格员,\n" +
"count(distinct case when\n" +
"rfus.is_paied = 1 and rfus.is_dtv = 1  then rfus.cust_id END) 数字电视缴费用户1,\n" +
"count(distinct case when\n" +
"rfus.is_paied = 1 and rfus.is_dtv = 1 then rfus.subscriber_ins_id  END) 数字电视缴费终端数2,\n" +
"count(distinct case when\n" +
"rfus.is_dbitv=1 and rfus.is_dbitv_paied = 1 then  rfus.cust_id END) 互动缴费用户数3,\n" +
"count(distinct case when\n" +
"rfus.is_dbitv=1 and rfus.is_dbitv_paied = 1 then rfus.subscriber_ins_id END) 互动缴费终端数4,\n" +
"count(distinct case when\n" +
"rfus.is_hdtv=1  and rfus.is_dbitv_paied = 1 and rfus.is_dbitv=1  then  rfus.cust_id END) 高清互动缴费用户数5,\n" +
"count(distinct case when\n" +
"rfus.is_hdtv=1  and rfus.is_dbitv_paied = 1 and rfus.is_dbitv=1  then  rfus.subscriber_ins_id END) 高清互动缴费终端数6,\n" +
"count(distinct case when\n" +
"rfus.is_lan_paied = 1  then rfus.cust_id END) 宽带缴费用户数7,\n" +
"count(distinct case when\n" +
"rfus.is_lan_paied = 1  then rfus.subscriber_ins_id END) 宽带缴费终端数8,\n" +
"count(distinct case when\n" +
"rfus.is_valid1 = 1 and rfus.is_dtv = 1 then rfus.cust_id END) 数字电视有效用户数9,\n" +
"count(distinct case when\n" +
"rfus.is_valid1 = 1 and rfus.is_dtv = 1 then rfus.subscriber_ins_id END) 数字电视有效终端数10,-- 李诚：is_valid1好像是欠费停机1年以内的，2是两年\n" +
"count(distinct case when cust.cust_type=1   --住宅\n" +
"and rfus.is_valid1 = 1 and rfus.is_dtv = 1 then rfus.cust_id END) 住宅数字电视有效用户数11,\n" +
"count(distinct case when\n" +
"rfus.is_dbitv=1 and rfus.is_valid1 = 1 then rfus.cust_id END) 互动有效用户数12,\n" +
"count(distinct case when\n" +
"rfus.is_dbitv=1 and rfus.is_valid1 = 1 then rfus.subscriber_ins_id END) 互动有效终端数13,\n" +
"count(distinct case when\n" +
"rfus.is_dbitv=1 and rfus.is_valid1= 1 and rfus.is_hdtv = 1   then rfus.cust_id END) 高清互动有效用户数14,\n" +
"count(distinct case when\n" +
"rfus.is_dbitv=1 and rfus.is_valid1= 1 and rfus.is_hdtv = 1 then rfus.subscriber_ins_id END) 高清互动有效终端数15,\n" +
"count(distinct case when\n" +
"rfus.is_lan_valid = 1 then rfus.cust_id END) 宽带有效用户数16,\n" +
"count(distinct case when\n" +
"rfus.is_lan_valid = 1 then rfus.subscriber_ins_id END) 宽带有效终端数17,\n" +
"count(distinct case when  cust.cust_type <> 1   --非住宅\n" +
"and rfus.is_paied = 1 and rfus.is_dtv = 1\n" +
"and rfus.user_name not like '%test%'\n" +
"then rfus.subscriber_ins_id  END) 非住宅数字电视缴费终端数18\n" +
"from  rep2.rep_fact_um_subscriber_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddDays(-1).ToString("yyyyMMdd") + "  rfus, cp2.cm_customer cust, wxjy.jy_customer_wg_rel wg, wxjy.jy_region_address_rel rad\n" +
"where cust.cust_id = rfus.cust_id and cust.cust_id = wg.cust_id(+) and cust.cust_id = rad.cust_id(+)\n" +
"and rfus.corp_org_id = 3328\n" +
"group by wg.grid_name, wg.mgr_name, rad.jy_region_name\n" +
"order by wg.grid_name, wg.mgr_name, rad.jy_region_name";
                int[] columntxt = { 0 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\日报\\" + "日报云BOSS宽表综合统计带网格人员--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 日报全量100m及以上的宽带明细数据
        public static bool baobiao17()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "日报全量100m及以上的宽带明细数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                subs.login_name 宽带登录名,\n" +
"                fsta.create_date 产品订购时间,\n" +
"                prod.offer_name 产品,\n" +
"                subs.create_date 终端开户时间,\n" +
"                case\n" +
"                  when subs.state = '1' and fsta.offer_status = '1' and\n" +
"                       fsta.os_status is null then\n" +
"                   '正常'\n" +
"                  when subs.state = 'M' then\n" +
"                   '暂停'\n" +
"                  when subs.state = 'D' then\n" +
"                   '注销'\n" +
"                  when subs.state = '1' then\n" +
"                   decode(fsta.os_status,\n" +
"                          1,\n" +
"                          '欠费停',\n" +
"                          2,\n" +
"                          '欠费连带停',\n" +
"                          3,\n" +
"                          '主动停',\n" +
"                          4,\n" +
"                          '管理停',\n" +
"                          5,\n" +
"                          '使用费不足停',\n" +
"                          6,\n" +
"                          '剪线停')\n" +
"                  else\n" +
"                   ''\n" +
"                end 停开机状态,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_address          addr,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_offer_06         ofer,\n" +
"       files2.um_offer_sta_02     fsta,\n" +
"       upc1.pm_offer              prod,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_customer_wg_rel    wg\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"   and fsta.offer_id = prod.offer_id\n" +
"   and subs.login_name is not null\n" +
"   and ofer.prod_service_id = 1004\n" +
"   and addr.expire_date > sysdate\n" +
"   and fsta.expire_date > sysdate\n" +
"   and ofer.expire_date > sysdate\n" +
"   and (prod.offer_name like '%100M%' or prod.offer_name like '%200M%' or prod.offer_name like '%300M%' or\n" +
"       prod.offer_name like '%500M%'  or  prod.offer_name like '%600M%' or prod.offer_name like '%1000M%' or prod.offer_name like '%电视机%' or prod.offer_name like '%广电监控服务%' or prod.offer_name like '%宽带体验升级%')\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name,\n" +
"          addr.std_addr_name,\n" +
"          cust.cust_code,\n" +
"          subs.login_name";

                int[] columntxt = { 1, 4 }; //哪些列是文本格式
                int[] columndate = { 5, 7 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\日报\\" + "日报全量100m及以上的宽带明细数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 日报今年新增高清机顶盒明细
        public static bool baobiao18()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "日报今年新增高清机顶盒明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                term.serial_no 机顶盒,\n" +
"                subs.sub_bill_id 智能卡,\n" +
"                rsku.res_sku_name 资源类型,\n" +
"                ures.create_date 资源订购时间,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格\n" +
"  from cp2.cm_customer            cust,\n" +
"       files2.cm_account          acct,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_res              ures,\n" +
"       res1.res_terminal          term,\n" +
"       res1.res_sku               rsku,\n" +
"       files2.um_address          addr,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       cp2.cb_party               part,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_customer_wg_rel    wg\n" +
" where cust.cust_id = acct.cust_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"   and subs.subscriber_ins_id = ures.subscriber_ins_id\n" +
"   and ures.res_equ_no = term.serial_no\n" +
"   and term.res_sku_id = rsku.res_sku_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and addr.expire_date > sysdate\n" +
"   and ures.expire_date > sysdate\n" +
"   and rsku.res_type_id = 2\n" +
"   and rsku.res_sku_name in ('银河高清基本型HDC6910(江阴)',\n" +
"                             '银河高清交互型HDC691033(江阴)',\n" +
"                             '银河智能高清交互型HDC6910798(江阴)',\n" +
"                             '银河4K交互型HDC691090',\n" +
"                             '4K超高清型II型融合型（EOC）',\n" +
"                             '银河4K超高清II型融合型HDC6910B1(江阴)',\n" +
"                             '4K超高清简易型（基本型）')\n" +
"   and ures.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-01-01")).ToString("yyyy-MM-dd") + "'\n" +
"   and ures.create_date < date '" + DateTime.Now.ToString("yyyy-MM-dd") + "'\n" +
"   and cust.own_corp_org_id = 3328\n" +
"   and exists\n" +
" (select 1\n" +
"          from files2.um_offer_06 ofer, files2.um_offer_sta_02 fsta\n" +
"         where ofer.offer_ins_id = fsta.offer_ins_id\n" +
"           and ofer.prod_service_id = 1002\n" +
"           and fsta.offer_status = '1'\n" +
"           and fsta.os_status is null\n" +
"           and ofer.subscriber_ins_id = subs.subscriber_ins_id)\n" +
" order by cust.cust_code, term.serial_no";
                int[] columntxt = { 1, 4, 5, 8 }; //哪些列是文本格式
                int[] columndate = { 7 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\日报\\" + "日报今年新增高清机顶盒明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 日报2023年3月底没互动当前有开通互动的机顶盒明细
        public static bool baobiao19()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "日报2023年3月底没互动当前有开通互动的机顶盒明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                term.serial_no 机顶盒,\n" +
"                subs.sub_bill_id 智能卡,\n" +
"                subs.create_date 用户新装时间,\n" +
"                case\n" +
"                  when nvl(subs.main_subscriber_ins_id, 0) = 0 then\n" +
"                   '主机'\n" +
"                  else\n" +
"                   ''\n" +
"                end 主副机,\n" +
"                rsku.res_sku_name 资源类型,\n" +
"                ures.create_date 资源订购时间,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_res              ures,\n" +
"       res1.res_terminal          term,\n" +
"       res1.res_sku               rsku,\n" +
"       files2.um_address          addr,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_customer_wg_rel    wg\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ures.subscriber_ins_id\n" +
"   and ures.res_equ_no = term.serial_no\n" +
"   and term.res_sku_id = rsku.res_sku_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and addr.expire_date > sysdate\n" +
"   and ures.expire_date > sysdate\n" +
"   and rsku.res_type_id = 2\n" +
"   and rsku.res_sku_name in ('银河高清基本型HDC6910(江阴)',\n" +
"                             '银河高清交互型HDC691033(江阴)',\n" +
"                             '银河智能高清交互型HDC6910798(江阴)',\n" +
"                             '银河4K交互型HDC691090',\n" +
"                             '4K超高清型II型融合型（EOC）',\n" +
"                             '银河4K超高清II型融合型HDC6910B1(江阴)',\n" +
"                             '4K超高清简易型（基本型）')\n" +
"   and cust.own_corp_org_id = 3328\n" +
"   and exists\n" +
" (select 1\n" +
"          from files2.um_offer_06 ofer, files2.um_offer_sta_02 fsta\n" +
"         where ofer.offer_ins_id = fsta.offer_ins_id\n" +
"           and ofer.expire_date > sysdate\n" +
"           and fsta.expire_date > sysdate\n" +
"           and ofer.prod_service_id = 1002\n" +
"           and fsta.offer_status = '1'\n" +
"           and fsta.os_status is null\n" +
"           and ofer.subscriber_ins_id = subs.subscriber_ins_id) ---当前有正常的基本节目\n" +
"   and not exists\n" +
" (select 1\n" +
"          from rep2.rep_fact_ins_srvpkg_20230331 insp,\n" +
"               wxjy.jy_dbitv_product             jitv1\n" +
"         where insp.srvpkg_id = jitv1.offer_id\n" +
"           and insp.subscriber_ins_id = subs.subscriber_ins_id) ---2023年3月31号之前完全互动产品\n" +
"\n" +
"   and not exists (select 1\n" +
"          from rep2.rep_fact_um_subscriber_20230331 a\n" +
"         where a.is_dbitv = 1\n" +
"           and a.subscriber_ins_id = subs.subscriber_ins_id) ---2023年3月31号不算互动用户\n" +
"   and exists (select 1\n" +
"          from files2.um_offer_06     ofer2,\n" +
"               files2.um_offer_sta_02 fsta2,\n" +
"               wxjy.jy_dbitv_product  jitv2\n" +
"         where ofer2.offer_ins_id = fsta2.offer_ins_id\n" +
"           and fsta2.offer_id = jitv2.offer_id\n" +
"           and ofer2.expire_date > sysdate\n" +
"           and fsta2.expire_date > sysdate\n" +
"           and fsta2.create_date >= date '2023-4-1'\n" +
"           and fsta2.offer_status = '1'\n" +
"           and fsta2.os_status is null\n" +
"           and ofer2.subscriber_ins_id = subs.subscriber_ins_id) ---当前有正常互动产品\n" +
"   and not exists\n" +
" (select 1\n" +
"          from files2.um_offer_06     ofer3,\n" +
"               files2.um_offer_sta_02 fsta3,\n" +
"               wxjy.jy_dbitv_product  jitv3\n" +
"         where ofer3.offer_ins_id = fsta3.offer_ins_id\n" +
"           and fsta3.offer_id = jitv3.offer_id\n" +
"           and ofer3.expire_date > sysdate\n" +
"           and fsta3.expire_date > sysdate\n" +
"           and fsta3.create_date < date '2023-4-1'\n" +
"           and ofer3.subscriber_ins_id = subs.subscriber_ins_id) ---没有4月1号以前的互动产品\n" +
" order by rad.JY_REGION_NAME,\n" +
"          addr.std_addr_name,\n" +
"          cust.cust_code,\n" +
"          term.serial_no";

                int[] columntxt = { 1, 4, 5, 10 }; //哪些列是文本格式
                int[] columndate = { 6, 9 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\日报\\" + "日报2023年3月底没互动当前有开通互动的机顶盒明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 日报全量100M以上的包年宽带明细数据(带状态)
        public static bool baobiao20()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "日报全量100M以上的包年宽带明细数据(带状态)--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                subs.login_name 宽带登录名,\n" +
"                prod.offer_name 产品,\n" +
"                ofer.create_date 订购时间,\n" +
"                ofer.valid_date 生效时间,\n" +
"                ofer.expire_date 失效时间,\n" +
"                case\n" +
"                  when subs.state = '1' and fsta.offer_status = '1' and\n" +
"                       fsta.os_status is null then\n" +
"                   '正常'\n" +
"                  when subs.state = 'M' then\n" +
"                   '暂停'\n" +
"                  when subs.state = 'D' then\n" +
"                   '注销'\n" +
"                  when fsta.offer_status = '3' and fsta.os_status = '1' then\n" +
"                   '欠费停'\n" +
"                  when fsta.offer_status = '3' and fsta.os_status = '3' then\n" +
"                   '暂停'\n" +
"                  else\n" +
"                   ''\n" +
"                end 停开机状态,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_address          addr,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_offer_06         ofer,\n" +
"       files2.um_offer_sta_02     fsta,\n" +
"       upc1.pm_offer              prod,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_customer_wg_rel    wg\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"   and fsta.offer_id = prod.offer_id\n" +
"   and ofer.expire_date > sysdate\n" +
"   and fsta.expire_date > sysdate\n" +
"   and addr.expire_date > sysdate\n" +
"   and subs.main_spec_id = 80020003\n" +
"   and subs.login_name is not null\n" +
"   and prod.offer_name in ('100M包年_300元_无锡地区',\n" +
"                           '100M包两年_600元_无锡地区',\n" +
"                           '300M包年_360元_无锡地区',\n" +
"                           '300M包两年_720元_无锡地区',\n" +
"                           '600M包年_480元_无锡地区',\n" +
"                           '600M包两年_960元_无锡地区',\n" +
"                           '1000M包年_600元_无锡地区')\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name,\n" +
"          addr.std_addr_name,\n" +
"          cust.cust_code,\n" +
"          ofer.create_date";
                int[] columntxt = { 1, 4, 10 }; //哪些列是文本格式
                int[] columndate = { 6, 7, 8 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\日报\\" + "日报全量100M以上的包年宽带明细数据(带状态)--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region FTTH缴费机顶盒明细数据
        public static bool baobiao21()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "FTTH缴费机顶盒明细数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                term.serial_no 机顶盒号码,\n" +
"                subs.sub_bill_id 智能卡,\n" +
"                rsku.res_sku_name 资源类型,\n" +
"                case\n" +
"                  when nvl(subs.main_subscriber_ins_id, 0) = 0 then\n" +
"                   '主机'\n" +
"                  else\n" +
"                   ''\n" +
"                end 主副机,\n" +
"                '正常' 停开机状态,\n" +
"                case\n" +
"                  when exists (select 1\n" +
"                          from files2.um_subscriber   subs1,\n" +
"                               files2.um_offer_06     ofer1,\n" +
"                               files2.um_offer_sta_02 fsta1\n" +
"                         where subs1.subscriber_ins_id =\n" +
"                               ofer1.subscriber_ins_id\n" +
"                           and ofer1.offer_ins_id = fsta1.offer_ins_id\n" +
"                           and subs1.login_name is not null\n" +
"                           and ofer1.prod_service_id = 1004\n" +
"                           and fsta1.offer_status = '1'\n" +
"                           and fsta1.os_status is null\n" +
"                           and subs1.main_spec_id = 80020003\n" +
"                           and ofer1.expire_date > sysdate\n" +
"                           and subs1.expire_date > sysdate\n" +
"                           and subs1.cust_id = cust.cust_id) then\n" +
"                   '有'\n" +
"                  else\n" +
"                   ''\n" +
"                end 是否有正常宽带,\n" +
"                case\n" +
"                  when exists (select 1\n" +
"                          from files2.um_offer_06     ofer2,\n" +
"                               files2.um_offer_sta_02 fsta2,\n" +
"                               upc1.pm_offer          prod2,\n" +
"                               wxjy.jy_dbitv_product  dbtv2\n" +
"                         where ofer2.offer_ins_id = fsta2.offer_ins_id\n" +
"                           and fsta2.offer_id = prod2.offer_id\n" +
"                           and prod2.offer_id = dbtv2.offer_id\n" +
"                           and fsta2.expire_date > sysdate\n" +
"                           and ofer2.expire_date > sysdate\n" +
"                           and ofer2.subscriber_ins_id =\n" +
"                               subs.subscriber_ins_id) then\n" +
"                   '有'\n" +
"                  else\n" +
"                   null\n" +
"                end 是否有互动,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_res              ures,\n" +
"       res1.res_terminal          term,\n" +
"       res1.res_sku               rsku,\n" +
"       files2.um_offer_06         ofer,\n" +
"       files2.um_offer_sta_02     fsta,\n" +
"       files2.um_address          addr,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_customer_wg_rel    wg\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"   and subs.subscriber_ins_id = ures.subscriber_ins_id\n" +
"   and ures.res_equ_no = term.serial_no\n" +
"   and term.res_sku_id = rsku.res_sku_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and cust.party_id = rel.party_id(+)\n" +
"   and cust.partition_id = rel.partition_id(+)\n" +
"   and rsku.res_type_id = 2\n" +
"   and ofer.prod_service_id = 1002\n" +
"   and fsta.offer_status = '1'\n" +
"   and fsta.os_status is null\n" +
"   and ures.expire_date > sysdate\n" +
"   and addr.expire_date > sysdate\n" +
"   and fsta.expire_date > sysdate\n" +
"   and ofer.expire_date > sysdate\n" +
"   and cust.own_corp_org_id = 3328\n" +
"   and not exists\n" +
" (select 1\n" +
"          from files2.um_subscriber   subs1,\n" +
"               files2.um_offer_06     ofer1,\n" +
"               files2.um_offer_sta_02 fsta1,\n" +
"               upc1.pm_offer          prod1,\n" +
"               wxjy.jy_dbitv_product  dbtv1\n" +
"         where subs1.subscriber_ins_id = ofer1.subscriber_ins_id\n" +
"           and ofer1.offer_ins_id = fsta1.offer_ins_id\n" +
"           and fsta1.offer_id = prod1.offer_id\n" +
"           and prod1.offer_id = dbtv1.offer_id\n" +
"           and subs1.main_spec_id = 80020199\n" +
"           and fsta1.expire_date > sysdate\n" +
"           and ofer1.expire_date > sysdate\n" +
"           and subs1.cust_id = cust.cust_id)\n" +
" order by rad.jy_region_name,\n" +
"          addr.std_addr_name,\n" +
"          cust.cust_code,\n" +
"          term.serial_no";

                int[] columntxt = { 1, 4,5,11 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\日报\\" + "FTTH缴费机顶盒明细数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region FTTH缴费宽带用户明细数据
        public static bool baobiao22()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "FTTH缴费宽带用户明细数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                          case\n" +
"                  when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or \n" +
"                        (cust.cust_type = 4 and cust.cust_prop = 10)) then\n" +
"                   '业务用机(免催免停)'\n" +
"                  else\n" +
"                   '普通'\n" +
"                end 客户属性,\n" +
"                subs.login_name 宽带登录名,\n" +
"                prod.offer_name 产品,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_address          addr,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_offer_06         ofer,\n" +
"       files2.um_offer_sta_02     fsta,\n" +
"       upc1.pm_offer              prod,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_customer_wg_rel    wg\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.party_id = rel.party_id(+)\n" +
"   and cust.partition_id = rel.partition_id(+)\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"   and fsta.offer_id = prod.offer_id\n" +
"   and ofer.prod_service_id = 1004\n" +
"   and fsta.offer_status = '1'\n" +
"   and fsta.os_status is null\n" +
"   and subs.login_name is not null\n" +
"   and addr.expire_date > sysdate\n" +
"   and ofer.expire_date > sysdate\n" +
"   and fsta.expire_date > sysdate\n" +
"   and cust.corp_org_id = 3328\n" +
" order by addr.std_addr_name, cust.cust_code, subs.login_name";
                int[] columntxt = { 1, 5, 7 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\日报\\" + "FTTH缴费宽带用户明细数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region FTTH纯标清客户的缴费机顶盒明细
        public static bool baobiao23()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "FTTH纯标清客户的缴费机顶盒明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                decode(rela.state, 0, '是', 1, '') 是否银行代扣,\n" +
"                term.serial_no 机顶盒号码,\n" +
"                subs.sub_bill_id 智能卡,\n" +
"                rsku.res_sku_name 资源类型,\n" +
"                case\n" +
"                  when nvl(subs.main_subscriber_ins_id, 0) = 0 then\n" +
"                   '主机'\n" +
"                  else\n" +
"                   ''\n" +
"                end 主副机,\n" +
"                case\n" +
"                  when subs.state = '1' and fsta.offer_status = '1' and\n" +
"                       fsta.os_status is null then\n" +
"                   '正常'\n" +
"                  when subs.state = 'M' then\n" +
"                   '暂停'\n" +
"                  when subs.state = 'D' then\n" +
"                   '注销'\n" +
"                  when fsta.offer_status = '3' and fsta.os_status = '1' then\n" +
"                   '欠费停'\n" +
"                  when fsta.offer_status = '3' and fsta.os_status = '3' then\n" +
"                   '暂停'\n" +
"                  else\n" +
"                   ''\n" +
"                end 停开机状态,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格\n" +
"  from cp2.cm_customer            cust,\n" +
"       files2.cm_account          acct,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_res              ures,\n" +
"       res1.res_terminal          term,\n" +
"       res1.res_sku               rsku,\n" +
"       files2.um_offer_06         ofer,\n" +
"       files2.um_offer_sta_02     fsta,\n" +
"       files2.um_address          addr,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_customer_wg_rel    wg,\n" +
"       ac2.am_entrust_relation    rela\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = acct.cust_id\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and acct.acct_id = rela.acct_id(+)\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"   and subs.subscriber_ins_id = ures.subscriber_ins_id\n" +
"   and ures.res_equ_no = term.serial_no\n" +
"   and term.res_sku_id = rsku.res_sku_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and rsku.res_type_id = 2\n" +
"   and ofer.prod_service_id = 1002\n" +
"   and ures.expire_date > sysdate\n" +
"   and addr.expire_date > sysdate\n" +
"   and fsta.expire_date > sysdate\n" +
"   and ofer.expire_date > sysdate\n" +
"   and cust.own_corp_org_id = 3328\n" +
"   and not exists\n" +
" (select 1\n" +
"          from files2.um_subscriber subs1,\n" +
"               files2.um_res        ures1,\n" +
"               res1.res_terminal    term1,\n" +
"               res1.res_sku         rsku1\n" +
"         where subs1.subscriber_ins_id = ures1.subscriber_ins_id\n" +
"           and ures1.res_equ_no = term1.serial_no\n" +
"           and term1.res_sku_id = rsku1.res_sku_id\n" +
"           and ures1.res_type_id = 2\n" +
"           and rsku1.res_sku_name in ('银河高清基本型HDC6910(江阴)',\n" +
"                                      '银河高清交互型HDC691033(江阴)',\n" +
"                                      '银河智能高清交互型HDC6910798(江阴)',\n" +
"                                      '银河4K交互型HDC691090',\n" +
"                                      '4K超高清型II型融合型（EOC）',\n" +
"                                      '银河4K超高清II型融合型HDC6910B1(江阴)',\n" +
"                                      '4K超高清简易型（基本型）')\n" +
"           and subs1.cust_id = cust.cust_id)\n" +
"   and exists\n" +
" (select 1\n" +
"          from files2.um_offer_06 ofer2, files2.um_offer_sta_02 fsta2\n" +
"         where ofer2.offer_ins_id = fsta2.offer_ins_id\n" +
"           and fsta2.expire_date > sysdate\n" +
"           and ofer2.expire_date > sysdate\n" +
"           and fsta2.offer_status = '1'\n" +
"           and ofer2.prod_service_id = 1002\n" +
"           and fsta2.os_status is null\n" +
"           and ofer2.subscriber_ins_id = subs.subscriber_ins_id) --基本包为开通\n" +
" order by rad.jy_region_name,\n" +
"          addr.std_addr_name,\n" +
"          cust.cust_code,\n" +
"          term.serial_no";
                int[] columntxt = { 1, 5, 6,10 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\日报\\" + "FTTH纯标清客户的缴费机顶盒明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region FTTH欠费停机以及暂停用户明细
        public static bool baobiao24()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "FTTH欠费停机以及暂停用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                term.serial_no 机顶盒,\n" +
"                subs.sub_bill_id 智能卡,\n" +
"                case\n" +
"                  when nvl(subs.main_subscriber_ins_id, 0) = 0 then\n" +
"                   '主机'\n" +
"                  else\n" +
"                   ''\n" +
"                end 主副机,\n" +
"                case\n" +
"                  when subs.state = 'M' then\n" +
"                   '暂停'\n" +
"                  when fsta.os_status = '1' then\n" +
"                   '欠费停机'\n" +
"                  when fsta.os_status in ('3', '4') then\n" +
"                   '暂停'\n" +
"                  else\n" +
"                   ''\n" +
"                end 停开机状态,\n" +
"                case\n" +
"                  when subs.state = 'M' then\n" +
"                   subs.done_date\n" +
"                  when fsta.os_status = '1' then\n" +
"                   fsta.done_date\n" +
"                  when fsta.os_status in (3, 4) then\n" +
"                   subs.done_date\n" +
"                  else\n" +
"                   null\n" +
"                end 受理时间,\n" +
"                rsku.res_sku_name 资源类型,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_offer_06         ofer,\n" +
"       files2.um_offer_sta_02     fsta,\n" +
"       files2.um_res              ures,\n" +
"       res1.res_terminal          term,\n" +
"       res1.res_sku               rsku,\n" +
"       files2.um_address          addr,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_customer_wg_rel    wg\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ures.subscriber_ins_id\n" +
"   and ures.res_equ_no = term.serial_no\n" +
"   and term.res_sku_id = rsku.res_sku_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"   and ofer.prod_service_id = 1002 --- 基本节目\n" +
"   and subs.main_spec_id = 80020001 ---数字电视\n" +
"   and addr.expire_date > sysdate\n" +
"   and ures.expire_date > sysdate\n" +
"   and fsta.expire_date > sysdate\n" +
"   and ofer.expire_date > sysdate\n" +
"   and rsku.res_type_id = 2\n" +
"   and cust.own_corp_org_id = 3328\n" +
"   and not exists\n" +
" (select 1\n" +
"          from files2.um_offer_06 ofer1, files2.um_offer_sta_02 fsta1\n" +
"         where ofer1.offer_ins_id = fsta1.offer_ins_id\n" +
"           and ofer1.prod_service_id = 1002\n" +
"           and fsta1.offer_status = '1'\n" +
"           and fsta1.os_status is null\n" +
"           and fsta1.expire_date > sysdate\n" +
"           and ofer1.expire_date > sysdate\n" +
"           and ofer1.subscriber_ins_id = subs.subscriber_ins_id)\n" +
" order by rad.jy_region_name,\n" +
"          addr.std_addr_name,\n" +
"          cust.cust_code,\n" +
"          term.serial_no";
                int[] columntxt = { 1, 4, 5, 10 }; //哪些列是文本格式
                int[] columndate = { 8 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\日报\\" + "FTTH欠费停机以及暂停用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 日报6月30号缴费机顶盒没有互动当前订购互动的明细
        public static bool baobiao25()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "6月30号缴费机顶盒没有互动当前订购互动的明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                term.serial_no 机顶盒,\n" +
"                subs.sub_bill_id 智能卡,\n" +
"                rsku.res_sku_name 资源型号,\n" +
"                prod1.offer_name 套餐,\n" +
"                prod.offer_name 产品,\n" +
"                fsta.create_date 订购时间,\n" +
"                fsta.valid_date 生效时间,\n" +
"                fsta.expire_date 失效时间,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_address          addr,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_res              ures,\n" +
"       res1.res_terminal          term,\n" +
"       res1.res_sku               rsku,\n" +
"       files2.um_offer_06         ofer,\n" +
"       files2.um_offer_sta_02     fsta,\n" +
"       upc1.pm_offer              prod,\n" +
"       upc1.pm_offer              prod1,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_customer_wg_rel    wg\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ures.subscriber_ins_id\n" +
"   and ures.res_equ_no = term.serial_no\n" +
"   and term.res_sku_id = rsku.res_sku_id\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"   and fsta.offer_id = prod.offer_id\n" +
"   and fsta.parent_offer_id = prod1.offer_id\n" +
"   and rsku.res_type_id = 2\n" +
"   and fsta.offer_status = '1'\n" +
"   and fsta.os_status is null\n" +
"   and addr.expire_date > sysdate\n" +
"   and ures.expire_date > sysdate\n" +
"   and fsta.expire_date > sysdate\n" +
"   and ofer.expire_date > sysdate\n" +
"   and fsta.create_date > date '2023-7-1'\n" +
"   and (prod.offer_name like '%芒果TV首年288元_江阴%' or\n" +
"       prod.offer_name like '%4K套餐B首年480元_江阴%' or\n" +
"       prod.offer_name like '%极视影院_4K套餐包年300元_江阴%' or\n" +
"       prod.offer_name like '%4K套餐C首年480元_江阴%' or\n" +
"       prod.offer_name like '%4K套餐E首年360元_江阴%' or\n" +
"       prod.offer_name like '%FTTH宽带升级购机套餐_江阴(首月96元)%' or\n" +
"       prod.offer_name like '%升级100_首年4元%' or\n" +
"       prod.offer_name like '%升级120套餐%' or\n" +
"       prod.offer_name like '%基本互动优惠产品_192元/24个月_江阴%')\n" +
"   and cust.own_corp_org_id = 3328\n" +
"   and exists\n" +
" (select 1\n" +
"          from rep2.rep_fact_um_subscriber_20230630 s1\n" +
"         where s1.is_dtv = 1\n" +
"           and s1.is_paied = 1\n" +
"           and not exists\n" +
"         (select 1\n" +
"                  from rep2.rep_fact_ins_srvpkg_20230630 s2,\n" +
"                       wxjy.jy_dbitv_product             p\n" +
"                 where s2.srvpkg_id = p.offer_id\n" +
"                   and s2.subscriber_ins_id = s1.subscriber_ins_id)\n" +
"           and s1.subscriber_ins_id = subs.subscriber_ins_id)\n" +
" order by rad.jy_region_name,\n" +
"          addr.std_addr_name,\n" +
"          cust.cust_code,\n" +
"          term.serial_no";
                int[] columntxt = { 1, 4, 5,10 }; //哪些列是文本格式
                int[] columndate = { 9,10,11 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\日报\\" + "6月30号缴费机顶盒没有互动当前订购互动的明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 日报6月30号已停机顶盒当前订购相关互动产品的明细
        public static bool baobiao26()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "6月30号已停机顶盒当前订购相关互动产品的明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                term.serial_no 机顶盒,\n" +
"                subs.sub_bill_id 智能卡,\n" +
"                rsku.res_sku_name 资源型号,\n" +
"                prod1.offer_name 套餐,\n" +
"                prod.offer_name 产品,\n" +
"                fsta.create_date 订购时间,\n" +
"                fsta.valid_date 生效时间,\n" +
"                fsta.expire_date 失效时间,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_address          addr,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_res              ures,\n" +
"       res1.res_terminal          term,\n" +
"       res1.res_sku               rsku,\n" +
"       files2.um_offer_06         ofer,\n" +
"       files2.um_offer_sta_02     fsta,\n" +
"       upc1.pm_offer              prod,\n" +
"       upc1.pm_offer              prod1,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_customer_wg_rel    wg\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ures.subscriber_ins_id\n" +
"   and ures.res_equ_no = term.serial_no\n" +
"   and term.res_sku_id = rsku.res_sku_id\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"   and fsta.offer_id = prod.offer_id\n" +
"   and fsta.parent_offer_id = prod1.offer_id\n" +
"   and rsku.res_type_id = 2\n" +
"   and fsta.offer_status = '1'\n" +
"   and fsta.os_status is null\n" +
"   and addr.expire_date > sysdate\n" +
"   and ures.expire_date > sysdate\n" +
"   and fsta.expire_date > sysdate\n" +
"   and ofer.expire_date > sysdate\n" +
"   and fsta.create_date > date '2023-7-1'\n" +
"   and (prod.offer_name like '%芒果TV首年288元_江阴%' or\n" +
"       prod.offer_name like '%4K套餐B首年480元_江阴%' or\n" +
"       prod.offer_name like '%极视影院_4K套餐包年300元_江阴%' or\n" +
"       prod.offer_name like '%4K套餐C首年480元_江阴%' or\n" +
"       prod.offer_name like '%4K套餐E首年360元_江阴%' or\n" +
"       prod.offer_name like '%FTTH宽带升级购机套餐_江阴(首月96元)%' or\n" +
"       prod.offer_name like '%升级100_首年4元%' or\n" +
"       prod.offer_name like '%升级120套餐%' or\n" +
"       prod.offer_name like '%基本互动优惠产品_192元/24个月_江阴%')\n" +
"   and cust.own_corp_org_id = 3328\n" +
"   and exists (select 1\n" +
"          from rep2.rep_fact_um_subscriber_20230630 s1\n" +
"         where s1.is_dtv = 1\n" +
"           and s1.subscriber_ins_id = subs.subscriber_ins_id)\n" +
"   and not exists\n" +
" (select 1\n" +
"          from rep2.rep_fact_um_subscriber_20230630 s2\n" +
"         where s2.is_dtv = 1\n" +
"           and s2.is_paied = 1\n" +
"           and s2.subscriber_ins_id = subs.subscriber_ins_id)\n" +
" order by rad.jy_region_name,\n" +
"          addr.std_addr_name,\n" +
"          cust.cust_code,\n" +
"          term.serial_no";




                int[] columntxt = { 1, 4, 5 }; //哪些列是文本格式
                int[] columndate = { 9,10,11 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\日报\\" + "6月30号已停机顶盒当前订购相关互动产品的明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 日报自7月份开始的新装机顶盒订购相关互动产品明细
        public static bool baobiao27()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\日报\\" + "自7月份开始的新装机顶盒订购相关互动产品明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                term.serial_no 机顶盒,\n" +
"                subs.sub_bill_id 智能卡,\n" +
"                rsku.res_sku_name 资源型号,\n" +
"                prod1.offer_name 套餐,\n" +
"                prod.offer_name 产品,\n" +
"                fsta.create_date 订购时间,\n" +
"                fsta.valid_date 生效时间,\n" +
"                fsta.expire_date 失效时间,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_address          addr,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_res              ures,\n" +
"       res1.res_terminal          term,\n" +
"       res1.res_sku               rsku,\n" +
"       files2.um_offer_06         ofer,\n" +
"       files2.um_offer_sta_02     fsta,\n" +
"       upc1.pm_offer              prod,\n" +
"       upc1.pm_offer              prod1,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_customer_wg_rel    wg\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ures.subscriber_ins_id\n" +
"   and ures.res_equ_no = term.serial_no\n" +
"   and term.res_sku_id = rsku.res_sku_id\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"   and fsta.offer_id = prod.offer_id\n" +
"   and fsta.parent_offer_id = prod1.offer_id\n" +
"   and rsku.res_type_id = 2\n" +
"   and fsta.offer_status = '1'\n" +
"   and fsta.os_status is null\n" +
"   and addr.expire_date > sysdate\n" +
"   and ures.expire_date > sysdate\n" +
"   and fsta.expire_date > sysdate\n" +
"   and ofer.expire_date > sysdate\n" +
"   and subs.create_date > date '2023-7-1'\n" +
"   and fsta.create_date > date '2023-7-1'\n" +
"   and (prod.offer_name like '%芒果TV首年288元_江阴%' or\n" +
"       prod.offer_name like '%4K套餐B首年480元_江阴%' or\n" +
"       prod.offer_name like '%极视影院_4K套餐包年300元_江阴%' or\n" +
"       prod.offer_name like '%4K套餐C首年480元_江阴%' or\n" +
"       prod.offer_name like '%4K套餐E首年360元_江阴%' or\n" +
"       prod.offer_name like '%FTTH宽带升级购机套餐_江阴(首月96元)%' or\n" +
"       prod.offer_name like '%升级100_首年4元%' or\n" +
"       prod.offer_name like '%升级120套餐%' or\n" +
"       prod.offer_name like '%基本互动优惠产品_192元/24个月_江阴%')\n" +
"   and cust.own_corp_org_id = 3328\n" +
"   and not exists\n" +
" (select 1\n" +
"          from rep2.rep_fact_um_subscriber_20230630 s1\n" +
"         where s1.is_dtv = 1\n" +
"           and s1.subscriber_ins_id = subs.subscriber_ins_id)\n" +
" order by rad.jy_region_name,\n" +
"          addr.std_addr_name,\n" +
"          cust.cust_code,\n" +
"          term.serial_no";
                int[] columntxt = { 1, 4, 5 }; //哪些列是文本格式
                int[] columndate = { 9, 10, 11 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\日报\\" + "自7月份开始的新装机顶盒订购相关互动产品明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion



        //        #region 周报数字客户级产品订购
        //        public static bool zbaobiao1()
        //        {
        //            string sqlString =

        //"select distinct cust.cust_code 客户证号,\n" +
        //"                        part.party_name 客户姓名,\n" +
        //"                        decode(cust.cust_type,\n" +
        //"                               1,\n" +
        //"                               '公众客户',\n" +
        //"                               2,\n" +
        //"                               '商业客户',\n" +
        //"                               3,\n" +
        //"                               '团体代付客户',\n" +
        //"                               4,\n" +
        //"                               '合同商业客户') 客户类型,\n" +
        //"                        case when cust.cust_type = 1 and cust.cust_prop = 6 then '测试客户' else '' end  只判断公众是否测试,\n" +
        //"                        '' 机顶盒,\n" +
        //"                        '' 智能卡,\n" +
        //"                        '' 资源型号,\n" +
        //"                        prod.offer_name 产品,\n" +
        //"                        '是' 客户级套餐产品,\n" +
        //"                        fsta.create_date 订购时间,\n" +
        //"                        fsta.valid_date 生效时间,\n" +
        //"                        staf.staff_name 操作员,\n" +
        //"                        orga.organize_name 营业员,\n" +
        //"                        rel.teleph_nunber 联系方式,\n" +
        //"                        rad.jy_region_name 区域,\n" +
        //"                        addr.std_addr_name 地址\n" +
        //"          from cp2.cm_customer            cust,\n" +
        //"               files2.cm_account          acct,\n" +
        //"               files2.um_subscriber       subs,\n" +
        //"               files2.um_offer_06         ofer,\n" +
        //"               files2.um_offer_sta_02     fsta,\n" +
        //"               upc1.pm_offer              prod,\n" +
        //"               wxjy.jy_region_address_rel rad,\n" +
        //"               wxjy.jy_contact_rel        rel,\n" +
        //"               files2.um_address          addr,\n" +
        //"               params1.sec_operator       oper,\n" +
        //"               params1.sec_staff          staf,\n" +
        //"       cp2.cb_party              part,\n" +
        //"               params1.sec_organize       orga\n" +
        //"         where cust.cust_id = subs.cust_id\n" +
        //"           and cust.cust_id = acct.acct_id\n" +
        //"           and cust.cust_id = rad.cust_id\n" +
        //"   and cust.party_id = part.party_id\n" +
        //"           and cust.cust_id = addr.cust_id\n" +
        //"           and cust.party_id = rel.party_id\n" +
        //"           and cust.partition_id = rel.partition_id\n" +
        //"           and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
        //"           and ofer.offer_ins_id = fsta.offer_ins_id\n" +
        //"           and fsta.offer_id = prod.offer_id\n" +
        //"           and ofer.op_id = oper.operator_id(+)\n" +
        //"           and oper.staff_id = staf.staff_id(+)\n" +
        //"           and staf.organize_id = orga.organize_id(+)\n" +
        //"           and subs.main_spec_id = 80020199\n" +
        //"           and fsta.offer_status = '1'\n" +
        //"           and fsta.os_status is null\n" +
        //"           and  fsta.valid_date >= trunc(sysdate)-7\n" +
        //"           and  fsta.valid_date < trunc(sysdate) \n" +
        //"           and (prod.offer_name like '%奇异%' or prod.offer_name like '%奇艺%' or\n" +
        //"               prod.offer_name like '%芒果%' or\n" +
        //"               prod.offer_name like '%学霸宝盒%' or\n" +
        //"               prod.offer_name like '%电竞视频%' or\n" +

        //"               prod.offer_name like '%致敬经典%' or\n" +
        //"               prod.offer_name like '%探奇动物界%' or\n" +
        //"               prod.offer_name like '%果果乐园%' or\n" +
        //"               prod.offer_name like '%上文广%')\n" +
        //"           and fsta.expire_date > sysdate\n" +
        //"           and cust.own_corp_org_id = 3328";
        //            int[] columntxt = { 1, 14 }; //哪些列是文本格式
        //            int[] columndate = { 10,11 };         //哪些列是日期格式
        //            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
        //            ExcelHelper.DataTableToExcel("\\周报\\" + "周报数字客户级产品订购--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
        //            return true;
        //        }
        //        #endregion

        //        #region 周报数字用户级产品订购
        //        public static bool zbaobiao2()
        //        {
        //            string sqlString =


        //"select distinct cust.cust_code 客户证号,\n" +
        //"                part.party_name 客户姓名,\n" +
        //"                decode(cust.cust_type,\n" +
        //"                       1,\n" +
        //"                       '公众客户',\n" +
        //"                       2,\n" +
        //"                       '商业客户',\n" +
        //"                       3,\n" +
        //"                       '团体代付客户',\n" +
        //"                       4,\n" +
        //"                       '合同商业客户') 客户类型,\n" +
        //"                term.serial_no 机顶盒,\n" +
        //"                subs.sub_bill_id 智能卡,\n" +
        //"                rsku.res_sku_name 资源型号,\n" +
        //"                prod.offer_name 产品,\n" +
        //"                '' 客户级套餐产品,\n" +
        //"                fsta.create_date 订购时间,\n" +
        //"                fsta.valid_date 生效时间,\n" +
        //"                fsta.expire_date 失效时间,\n" +
        //"                staf.staff_name 操作员,\n" +
        //"                orga.organize_name 营业厅,\n" +
        //"                rel.teleph_nunber 联系方式,\n" +
        //"                        case when cust.cust_type = 1 and cust.cust_prop =  '6' then '测试客户' else '' end  只判断公众是否测试,\n" +
        //"                rad.jy_region_name 区域,\n" +
        //"                addr.std_addr_name 地址\n" +
        //"  from cp2.cm_customer            cust,\n" +
        //"       files2.cm_account          acct,\n" +
        //"       files2.um_address          addr,\n" +
        //"       files2.um_subscriber       subs,\n" +
        //"       files2.um_res              ures,\n" +
        //"       res1.res_terminal          term,\n" +
        //"       res1.res_sku               rsku,\n" +
        //"       files2.um_offer_06         ofer,\n" +
        //"       files2.um_offer_sta_02     fsta,\n" +
        //"       upc1.pm_offer              prod,\n" +
        //"       wxjy.jy_contact_rel        rel,\n" +
        //"       wxjy.jy_region_address_rel rad,\n" +
        //"       params1.sec_operator       oper,\n" +
        //"       params1.sec_staff          staf,\n" +
        //"       cp2.cb_party              part,\n" +
        //"       params1.sec_organize       orga\n" +
        //" where cust.cust_id = acct.cust_id\n" +
        //"   and cust.cust_id = addr.cust_id\n" +
        //"   and cust.cust_id = rad.cust_id\n" +
        //"   and cust.party_id = part.party_id\n" +
        //"   and cust.party_id = rel.party_id\n" +
        //"   and cust.partition_id = rel.partition_id\n" +
        //"   and cust.cust_id = subs.cust_id\n" +
        //"   and subs.subscriber_ins_id = ures.subscriber_ins_id\n" +
        //"   and ures.res_equ_no = term.serial_no\n" +
        //"   and term.res_sku_id = rsku.res_sku_id\n" +
        //"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
        //"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
        //"   and fsta.offer_id = prod.offer_id\n" +
        //"   and ofer.op_id = oper.operator_id(+)\n" +
        //"   and oper.staff_id = staf.staff_id(+)\n" +
        //"   and staf.organize_id = orga.organize_id(+)\n" +
        //"   and rsku.res_type_id = 2\n" +
        //"   and fsta.offer_status = '1'\n" +
        //"   and fsta.os_status is null\n" +
        //"   and ures.expire_date > sysdate\n" +
        //"   and fsta.expire_date > sysdate\n" +
        //"   and fsta.valid_date >= trunc(sysdate) -7\n" +
        //"   and fsta.valid_date < trunc(sysdate)\n" +
        //"   and （prod.offer_name like '%奇异%' or prod.offer_name like '%奇艺%'  or prod.offer_name like '%炫力动漫%'   or prod.offer_name like '%芒果TV%'  or prod.offer_name like '%学霸宝盒%'\n" +
        //"    or prod.offer_name like '%电竞视频%'  or prod.offer_name like '%致敬经典%' or prod.offer_name like '%上文广%  or prod.offer_name like '%探奇动物界%' or prod.offer_name like '%果果乐园%' or prod.offer_name like '%电视机3年_0元/3年_江阴_首月出账%' or prod.offer_name like '%电视机5年_660元/3年_江阴_首月出账%'）\n" +
        //"   and cust.own_corp_org_id = 3328";

        //            int[] columntxt = { 1, 4, 5, 14 }; //哪些列是文本格式
        //            int[] columndate = { 9, 10,11 };         //哪些列是日期格式
        //            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
        //            ExcelHelper.DataTableToExcel("\\周报\\" + "周报数字用户级产品订购--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
        //            return true;
        //        }
        //        #endregion

        //        #region 周报数字产品退订
        //        public static bool zbaobiao3()
        //        {
        //            string sqlString =

        //"select distinct cust.cust_code 客户证号,\n" +
        //"                part.party_name 客户姓名,\n" +
        //"                decode(cust.cust_type,\n" +
        //"                       1,\n" +
        //"                       '公众客户',\n" +
        //"                       2,\n" +
        //"                       '商业客户',\n" +
        //"                       3,\n" +
        //"                       '团体代付客户',\n" +
        //"                       4,\n" +
        //"                       '合同商业客户') 客户类型,\n" +
        //"                       case when cust.cust_type = 1 and cust.cust_prop = 6 then '测试客户' else '' end  只判断公众是否测试,\n" +
        //"                subs.bill_id 机顶盒,\n" +
        //"                subs.sub_bill_id 智能卡,\n" +
        //"                prod.offer_name 套餐,\n" +
        //"                ofer.offer_name 产品,\n" +
        //"                ofer.create_date 订购时间,\n" +
        //"                ofer.valid_date 生效时间,\n" +
        //"                ofer.expire_date 失效时间,\n" +
        //"                staf.staff_name 操作员,\n" +
        //"                orga.organize_name 营业厅,\n" +
        //"                rel.teleph_nunber 联系方式,\n" +
        //"                rad.jy_region_name 区域,\n" +
        //"                addr.std_addr_name 地址\n" +
        //"  from cp2.cm_customer cust,\n" +
        //"       files2.cm_account acct,\n" +
        //"       cp2.cb_party              part,\n" +
        //"(select *\n" +
        //"   from JOUR2.Om_Subscriber\n" +
        //" union all\n" +
        //" select *\n" +
        //"   from JOUR2.Om_Subscriber_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "\n" +
        //"    union all\n" +
        //" select *\n" +
        //"   from JOUR2.Om_Subscriber_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).ToString("yyyyMM") + ") subs,\n" +
        //"(select *\n" +
        //"   from JOUR2.OM_OFFER\n" +
        //"    union all\n" +
        //" select *\n" +
        //"   from JOUR2.OM_OFFER_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "\n" +
        //"    union all\n" +
        //" select *\n" +
        //"   from JOUR2.OM_OFFER_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).ToString("yyyyMM") + ") ofer,\n" +
        //"(select *\n" +
        //"   from jour2.om_order\n" +
        //" union all\n" +
        //" select *\n" +
        //"   from jour2.om_order_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "\n" +
        //"    union all\n" +
        //" select *\n" +
        //"   from jour2.om_order_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).ToString("yyyyMM") + ") orde," +

        //"       params1.sec_operator oper,\n" +
        //"       params1.sec_staff staf,\n" +
        //"       params1.sec_organize orga,\n" +
        //"       wxjy.jy_contact_rel rel,\n" +
        //"       wxjy.jy_region_address_rel rad,\n" +
        //"       files2.um_address addr,\n" +
        //"       upc1.pm_offer prod\n" +
        //" where cust.cust_id = acct.cust_id\n" +
        //"   and cust.cust_id = rad.cust_id\n" +
        //"   and cust.party_id = part.party_id\n" +
        //"   and cust.cust_id = addr.cust_id\n" +
        //"   and cust.party_id = rel.party_id\n" +
        //"   and cust.partition_id = rel.partition_id\n" +
        //"   and cust.cust_id = orde.party_role_id\n" +
        //"   and orde.order_id = subs.order_id\n" +
        //"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
        //"   and ofer.parent_offer_id = prod.offer_id\n" +
        //"   and ofer.op_id = oper.operator_id(+)\n" +
        //"   and oper.staff_id = staf.staff_id(+)\n" +
        //"   and staf.organize_id = orga.organize_id(+)\n" +
        //"   and ofer.offer_type <> 'S'\n" +
        //"   and ofer.action = 1\n" +
        //"   and ofer.expire_date < sysdate\n" +
        //"   and orde.create_date >=  trunc(sysdate)-7 \n" +
        //"   and orde.create_date <  trunc(sysdate) \n" +
        //"   and (ofer.offer_name like '%奇异%' or ofer.offer_name like '%奇艺%' or ofer.offer_name like '%炫力动漫%' or ofer.offer_name like '%芒果TV%' or ofer.offer_name like '%学霸宝盒%' or ofer.offer_name like '%电竞视频%'\n" +
        //"    or ofer.offer_name like '%致敬经典%'  or ofer.offer_name like '%上文广%' or ofer.offer_name like '%探奇动物界%' or ofer.offer_name like '%果果乐园%' or ofer.offer_name like '%电视机3年_0元/3年_江阴_首月出账%' or ofer.offer_name like '%电视机5年_660元/3年_江阴_首月出账%')\n" +
        //"   and cust.own_corp_org_id = 3328\n" +
        //" order by cust.cust_code, subs.bill_id, rad.jy_region_name";
        //            int[] columntxt = { 1, 5,6,14 }; //哪些列是文本格式
        //            int[] columndate = { 9,10,11 };         //哪些列是日期格式
        //            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
        //            ExcelHelper.DataTableToExcel("\\周报\\" + "周报数字产品退订--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
        //            return true;
        //        }
        //        #endregion

        //        #region 周报FTTH宽带用户级产品订购
        //        public static bool zbaobiao4()
        //        {
        //            string sqlString =

        //"select distinct cust.cust_code 客户证号,\n" +
        //"                part.party_name 客户姓名,\n" +
        //"                decode(cust.cust_type,\n" +
        //"                       1,\n" +
        //"                       '公众客户',\n" +
        //"                       2,\n" +
        //"                       '商业客户',\n" +
        //"                       3,\n" +
        //"                       '团体代付客户',\n" +
        //"                       4,\n" +
        //"                       '合同商业客户') 客户类型,\n" +
        //"                subs.login_name 宽带登录名,\n" +
        //"                updf.offer_name 产品名称,\n" +
        //"                fsta.valid_date 生效时间,\n" +
        //"                 wgrel.grid_name 网格,\n" +
        //"                rad.jy_region_name 区域,\n" +
        //"                addr.std_addr_name 地址\n" +
        //"  from cp2.cm_customer            cust,\n" +
        //"       files2.cm_account          acct,\n" +
        //"       files2.um_subscriber       subs,\n" +
        //"       files2.um_offer_06         ofer,\n" +
        //"       files2.um_offer_sta_02     fsta,\n" +
        //"       upc1.pm_offer              updf,\n" +
        //"       wxjy.jy_customer_wg_rel              wgrel,\n" +
        //"        wxjy.jy_region_address_rel rad,\n" +
        //"       cp2.cb_party              part,\n" +
        //"         files2.um_address          addr\n" +
        //" where cust.cust_id = subs.cust_id\n" +
        //"   and cust.cust_id = acct.cust_id\n" +
        //"   and cust.party_id = part.party_id\n" +
        //"    and cust.cust_id = addr.cust_id\n" +
        //"   and cust.cust_id = rad.cust_id\n" +
        //"  and cust.cust_id=wgrel.cust_id\n" +
        //"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
        //"   and ofer.subscriber_ins_id = subs.subscriber_ins_id\n" +
        //"   and fsta.offer_id = updf.offer_id\n" +
        //"   and subs.state = '1'\n" +
        //"   and ofer.prod_service_id = 1004\n" +
        //"     and addr.expire_date > sysdate\n" +
        //"   and  fsta.expire_date > sysdate\n" +
        //"   and  fsta.valid_date >= trunc(sysdate)-7\n" +
        //"   and  fsta.valid_date <  trunc(sysdate)\n" +
        //" and (updf.offer_name like '%300M%' or updf.offer_name like '%500M%'  or updf.offer_name like '%600M%' or  updf.offer_name like '%1000M%' or  updf.offer_name like '%电视机%')\n" + 
        //"   and fsta.os_status is null\n" +
        //"   and cust.own_corp_org_id = 3328\n" +
        //" order by cust.cust_code";
        //            int[] columntxt = { 1, 4 }; //哪些列是文本格式
        //            int[] columndate = { 6 };         //哪些列是日期格式
        //            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
        //            ExcelHelper.DataTableToExcel("\\周报\\" + "周报FTTH宽带用户级产品订购--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
        //            return true;
        //        }
        //        #endregion

        //        #region 周报FTTH宽带产品退订
        //        public static bool zbaobiao5()
        //        {
        //            string sqlString =


        //"select distinct cust.cust_code 客户证号,\n" +
        //"                part.party_name 客户姓名,\n" +
        //"                decode(cust.cust_type,\n" +
        //"                       1,\n" +
        //"                       '公众客户',\n" +
        //"                       2,\n" +
        //"                       '商业客户',\n" +
        //"                       3,\n" +
        //"                       '团体代付客户',\n" +
        //"                       4,\n" +
        //"                       '合同商业客户') 客户类型,\n" +
        //"                       case when cust.cust_type = 1 and cust.cust_prop = 6 then '测试客户' else '' end  只判断公众是否测试,\n" +
        //"                prod.offer_name 套餐,\n" +
        //"                ofer.offer_name 产品,\n" +
        //"                ofer.create_date 订购时间,\n" +
        //"                ofer.valid_date 生效时间,\n" +
        //"                ofer.expire_date 失效时间,\n" +
        //"                staf.staff_name 操作员,\n" +
        //"                orga.organize_name 营业厅,\n" +
        //"                rel.teleph_nunber 联系方式,\n" +
        //"                rad.jy_region_name 区域,\n" +
        //"                addr.std_addr_name 地址\n" +
        //"  from cp2.cm_customer cust,\n" +
        //"       files2.cm_account acct,\n" +
        //"       cp2.cb_party              part,\n" +
        //"(select *\n" +
        //"   from JOUR2.Om_Subscriber\n" +
        //" union all\n" +
        //" select *\n" +
        //"   from JOUR2.Om_Subscriber_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "\n" +
        //"    union all\n" +
        //" select *\n" +
        //"   from JOUR2.Om_Subscriber_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).ToString("yyyyMM") + ") subs,\n" +
        //"(select *\n" +
        //"   from JOUR2.OM_OFFER\n" +
        //"    union all\n" +
        //" select *\n" +
        //"   from JOUR2.OM_OFFER_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "\n" +
        //"    union all\n" +
        //" select *\n" +
        //"   from JOUR2.OM_OFFER_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).ToString("yyyyMM") + ") ofer,\n" +
        //"(select *\n" +
        //"   from jour2.om_order\n" +
        //" union all\n" +
        //" select *\n" +
        //"   from jour2.om_order_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "\n" +
        //"    union all\n" +
        //" select *\n" +
        //"   from jour2.om_order_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).ToString("yyyyMM") + ") orde," +

        //"       params1.sec_operator oper,\n" +
        //"       params1.sec_staff staf,\n" +
        //"       params1.sec_organize orga,\n" +
        //"       wxjy.jy_contact_rel rel,\n" +
        //"       wxjy.jy_region_address_rel rad,\n" +
        //"       files2.um_address addr,\n" +
        //"       upc1.pm_offer prod\n" +
        //" where cust.cust_id = acct.cust_id\n" +
        //"   and cust.cust_id = rad.cust_id\n" +
        //"   and cust.party_id = part.party_id\n" +
        //"   and cust.cust_id = addr.cust_id\n" +
        //"   and cust.party_id = rel.party_id\n" +
        //"   and cust.partition_id = rel.partition_id\n" +
        //"   and cust.cust_id = orde.party_role_id\n" +
        //"   and orde.order_id = subs.order_id\n" +
        //"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
        //"   and ofer.parent_offer_id = prod.offer_id\n" +
        //"   and ofer.op_id = oper.operator_id(+)\n" +
        //"   and oper.staff_id = staf.staff_id(+)\n" +
        //"   and staf.organize_id = orga.organize_id(+)\n" +
        //"   and ofer.offer_type <> 'S'\n" +
        //"   and ofer.action = 1\n" +
        //"   and ofer.expire_date >=  trunc(sysdate)-7\n" +
        //"   and ofer.expire_date <  trunc(sysdate)\n" +
        //"   and (ofer.offer_name like '%300M%' or ofer.offer_name like '%500M%' or ofer.offer_name like '%600M%' or ofer.offer_name like '%1000M%'  or  ofer.offer_name like '%电视机%' )\n" +
        //"   and cust.own_corp_org_id = 3328";

        //            int[] columntxt = { 1, 12 }; //哪些列是文本格式
        //            int[] columndate = { 7,8,9 };         //哪些列是日期格式
        //            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
        //            ExcelHelper.DataTableToExcel("\\周报\\" + "周报FTTH宽带产品退订--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
        //            return true;
        //        }
        //        #endregion




        #region 月报各业务分账本的出账金额
        public static bool ybaobiao1()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "各业务分账本的出账金额--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select tp.service_id,\n" +
"       tp.service_name 分业务,\n" +
"       tp.jy_region_name 区域,\n" +
"       tp.cust_type 客户类型,\n" +
"       sum(a1) / 100 出账金额\n" +
"    /*   sum(a2) / 100 通用现金账本_销账,\n" +
"       sum(a3) / 100 数字基本现金账本_销账,\n" +
"       sum(a4) / 100 互动基本现金账本_销账,\n" +
"       sum(a5) / 100 专用账本_销账,\n" +
"       sum(a6) / 100 划转账本_销账,\n" +
"       sum(a7) / 100 虚拟通用账本_销账,\n" +
"       sum(a8) / 100 宽带现金账本_销账*/\n" +
"  from\n" +
"         ---分业务出账金额\n" +
"        (select sifo.service_id,\n" +
"                sifo.service_name,\n" +
"                rad.jy_region_name,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') cust_type,\n" +
"                nvl(sum(czmx.fee), 0) a1,   ---出账金额\n" +
"                0 a2,\n" +
"                0 a3,\n" +
"                0 a4,\n" +
"                0 a5,\n" +
"                0 a6,\n" +
"                0 a7,\n" +
"                0 a8\n" +
"           from cp2.cm_customer            cust,\n" +
"                wxjy.jy_region_address_rel rad,\n" +
"                files2.cm_account          acct,\n" +
"                ac2.am_bill_ar_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "      czmx,\n" +
"                pzg1.am_item_service       aits,\n" +
"                pzg1.am_service_info       sifo\n" +
"          where cust.cust_id = acct.cust_id\n" +
"            and cust.cust_id = rad.cust_id\n" +
"            and acct.acct_id = czmx.acct_id\n" +
"            and czmx.bill_item_id = aits.am_item_type_id\n" +
"            and aits.service_id = sifo.service_id\n" +
"            and cust.own_corp_org_id = 3328\n" +
"          group by sifo.service_id, sifo.service_name, rad.jy_region_name, cust.cust_type\n" +
"         union all\n" +
"         --分业务销账金额\n" +
"         select sifo.service_id,\n" +
"                sifo.service_name,\n" +
"                rad.jy_region_name,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') cust_type,\n" +
"                0 a1,\n" +
"                nvl(sum(case when a.asset_item_id = 100 then a.writeoff_fee else 0 end), 0) a2, --通用现金账本,\n" +
"                nvl(sum(case when a.asset_item_id = 200 then a.writeoff_fee else 0 end), 0) a3, --数字基本现金账本,\n" +
"                nvl(sum(case when a.asset_item_id = 300 then a.writeoff_fee else 0 end), 0) a4, --互动基本现金账本,\n" +
"                nvl(sum(case when a.asset_item_id = 400 then a.writeoff_fee else 0 end), 0) a5, --专用账本,\n" +
"                nvl(sum(case when a.asset_item_id = 500 then a.writeoff_fee else 0 end), 0) a6, --划转账本,\n" +
"                nvl(sum(case when a.asset_item_id = 600 then a.writeoff_fee else 0 end), 0) a7, --虚拟通用账本,\n" +
"                nvl(sum(case when a.asset_item_id = 700 then a.writeoff_fee else 0 end), 0) a8  --宽带现金账本\n" +
"           from (select cxz.acct_id,\n" +
"                        cxz.asset_item_id,\n" +
"                        cxz.bill_item_id,\n" +
"                        cxz.writeoff_fee,\n" +
"                        cxz.corp_org_id\n" +
"                   from ac2.am_writeoff cxz\n" +
"                  where cxz.cancel_flag = 'U'\n" +
"                    and cxz.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"                    and cxz.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"                    and cxz.corp_org_id = 3328\n" +
"                 union all\n" +
"                 select pxz.acct_id,\n" +
"                        pxz.asset_item_id,\n" +
"                        pxz.bill_item_id,\n" +
"                        pxz.writeoff_fee,\n" +
"                        pxz.corp_org_id\n" +
"                   from ac2.am_writeoff_d pxz\n" +
"                  where pxz.cancel_flag = 'U'\n" +
"                    and pxz.bill_month = " + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " \n" +
"                    and pxz.corp_org_id = 3328) a,\n" +
"                files2.cm_account acct,\n" +
"                cp2.cm_customer cust,\n" +
"                wxjy.jy_region_address_rel rad,\n" +
"                pzg1.am_asset_type asty,\n" +
"                pzg1.am_bill_type bilt,\n" +
"                pzg1.am_item_service aits,\n" +
"                pzg1.am_service_info sifo\n" +
"          where cust.cust_id = acct.cust_id\n" +
"            and cust.cust_id = rad.cust_id\n" +
"            and acct.acct_id = a.acct_id\n" +
"            and a.asset_item_id = asty.asset_item_id\n" +
"            and a.bill_item_id = bilt.bill_item_id\n" +
"            and bilt.bill_item_id = aits.am_item_type_id\n" +
"            and aits.service_id = sifo.service_id\n" +
"            and asty.asset_item_kind <= 12\n" +
"            and a.corp_org_id = 3328\n" +
"          group by rad.jy_region_name, cust.cust_type, sifo.service_id, sifo.service_name) tp\n" +
" group by tp.service_id, tp.service_name, tp.jy_region_name, tp.cust_type\n" +
" order by tp.jy_region_name, tp.service_id, tp.service_name, tp.cust_type";
                int[] columntxt = { 0 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "各业务分账本的出账金额--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报各业务账本余额
        public static bool ybaobiao2()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "各业务账本余额--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select aa.jy_region_name 区域,\n" +
"       aa.cust_type 客户类型,\n" +
"       aa.asset_item_name 账本,\n" +
"       aa.amount 销账前余额,\n" +
"       nvl(bb.writeoff_fee, 0) 批销金额,\n" +
"       aa.amount + nvl(bb.writeoff_fee, 0) 批销后余额\n" +
"  from (select rad.jy_region_name,\n" +
"               amty.asset_item_id,\n" +
"               amty.asset_item_name,\n" +
"               decode(cust.cust_type,\n" +
"                      1,\n" +
"                      '公众客户',\n" +
"                      2,\n" +
"                      '商业客户',\n" +
"                      3,\n" +
"                      '团体代付客户',\n" +
"                      4,\n" +
"                      '合同商业客户') cust_type,\n" +
"               sum(bala.balance) / 100 amount\n" +
"          from cp2.cm_customer                cust,\n" +
"               files2.cm_account              acct,\n" +
"               wxjy.jy_region_address_rel     rad,\n" +
"               rep2.rep_fact_balance_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " bala,\n" +
"               pzg1.am_asset_type             amty\n" +
"         where cust.cust_id = acct.cust_id\n" +
"           and cust.cust_id = rad.cust_id(+)\n" +
"           and acct.acct_id = bala.acct_id\n" +
"           and bala.asset_item_id = amty.asset_item_id\n" +
"           and cust.own_corp_org_id = 3328\n" +
"         group by rad.jy_region_name,\n" +
"                  amty.asset_item_name,\n" +
"                  cust.cust_type,\n" +
"                  amty.asset_item_id) aa,\n" +
"       (select rad.jy_region_name,\n" +
"               amty.asset_item_id,\n" +
"               amty.asset_item_name,\n" +
"               decode(cust.cust_type,\n" +
"                      1,\n" +
"                      '公众客户',\n" +
"                      2,\n" +
"                      '商业客户',\n" +
"                      3,\n" +
"                      '团体代付客户',\n" +
"                      4,\n" +
"                      '合同商业客户') cust_type,\n" +
"               -sum(ye.writeoff_fee) / 100 writeoff_fee\n" +
"          from cp2.cm_customer            cust,\n" +
"               files2.cm_account          acct,\n" +
"               wxjy.jy_region_address_rel rad,\n" +
"               ac2.am_writeoff_d          ye,\n" +
"               pzg1.am_asset_type         amty\n" +
"         where cust.cust_id = acct.cust_id\n" +
"           and cust.cust_id = rad.cust_id(+)\n" +
"           and acct.acct_id = ye.acct_id\n" +
"           and ye.asset_item_id = amty.asset_item_id\n" +
"           and ye.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"           and ye.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"           and cust.own_corp_org_id = 3328\n" +
"         group by rad.jy_region_name,\n" +
"                  amty.asset_item_name,\n" +
"                  cust.cust_type,\n" +
"                  amty.asset_item_id) bb\n" +
" where aa.jy_region_name = bb.jy_region_name(+)\n" +
"   and aa.asset_item_id = bb.asset_item_id(+)\n" +
"   and aa.cust_type = bb.cust_type(+)\n" +
" order by aa.jy_region_name, aa.asset_item_name, aa.cust_type";

                int[] columntxt = { 0 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "各业务账本余额--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报各业务月新增欠费金额
        public static bool ybaobiao3()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "各业务月新增欠费金额--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select rad.jy_region_name 区域,\n" +
"       sum(case when siof.service_id = 1002 then qfzd.balance else 0 end) / 100 数字基本业务,\n" +
"       sum(case when siof.service_id = 1003 then qfzd.balance else 0 end) / 100 互动基本业务,\n" +
"       sum(case when siof.service_id = 1004 then qfzd.balance else 0 end) / 100 宽带业务,\n" +
"       sum(case when siof.service_id = 1005 then qfzd.balance else 0 end) / 100 付费节目业务,\n" +
"       sum(case when siof.service_id = 1006 then qfzd.balance else 0 end) / 100 互动点播业务,\n" +
"       sum(case when siof.service_id = 1008 then qfzd.balance else 0 end) / 100 增值业务\n" +
"  from cp2.cm_customer      cust,\n" +
"       files2.cm_account    acct,\n" +
"       ac2.am_bill_item     qfzd,\n" +
"       pzg1.am_item_service aits,\n" +
"       pzg1.am_service_info siof,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where cust.cust_id = acct.cust_id\n" +
"   and cust.cust_id = rad.cust_id\n" +
"   and acct.acct_id = qfzd.acct_id\n" +
"   and qfzd.bill_item_id = aits.am_item_type_id\n" +
"   and aits.service_id = siof.service_id\n" +
"   and qfzd.bill_month = " + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "\n" +
"   and cust.own_corp_org_id = 3328\n" +
" group by rad.jy_region_name\n" +
" order by rad.jy_region_name";
                int[] columntxt = { 0 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "各业务月新增欠费金额--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报各月份各业务欠费总金额
        public static bool ybaobiao4()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "各月份各业务欠费总金额--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select qfzd.bill_month 月份,\n" +
"       sum(case when siof.service_id = 1002 then qfzd.balance else 0 end) / 100 数字基本业务,\n" +
"       sum(case when siof.service_id = 1003 then qfzd.balance else 0 end) / 100 互动基本业务,\n" +
"       sum(case when siof.service_id = 1004 then qfzd.balance else 0 end) / 100 宽带业务,\n" +
"       sum(case when siof.service_id = 1005 then qfzd.balance else 0 end) / 100 付费节目业务,\n" +
"       sum(case when siof.service_id = 1006 then qfzd.balance else 0 end) / 100 互动点播业务,\n" +
"       sum(case when siof.service_id = 1008 then qfzd.balance else 0 end) / 100 增值业务\n" +
"  from ac2.am_bill_item     qfzd,\n" +
"       pzg1.am_item_service aits,\n" +
"       pzg1.am_service_info siof\n" +
" where qfzd.bill_item_id = aits.am_item_type_id\n" +
"   and aits.service_id = siof.service_id\n" +
"   and qfzd.corp_org_id = 3328\n" +
" group by qfzd.bill_month\n" +
" order by qfzd.bill_month desc";
                int[] columntxt = { 0 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "各月份各业务欠费总金额--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报爱奇艺出销账(去测试客户)
        public static bool ybaobiao5()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "爱奇艺出销账(去测试客户)--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select cust.cust_code 客户证号,\n" +
"       part.party_name 姓名,\n" +
"       decode(cust.cust_type,\n" +
"              1,\n" +
"              '公众客户',\n" +
"              2,\n" +
"              '商业客户',\n" +
"              3,\n" +
"              '团体代付客户',\n" +
"              4,\n" +
"              '合同商业客户') 客户类型,\n" +
"       rad.jy_region_name 区域,\n" +
"       addr.std_addr_name 地址,\n" +
"       a1.amount1 出账金额,\n" +
"       a2.amont2 销账金额,\n" +
"                          case\n" +
"                  when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or \n" +
"                        (cust.cust_type = 4 and cust.cust_prop = 10)) then\n" +
"                   '业务用机(免催免停)'\n" +
"                  else\n" +
"                   '普通'\n" +
"                end 客户属性\n" +
"  from cp2.cm_customer cust,\n" +
"       files2.cm_account acct,\n" +
"       cp2.cb_party part,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       files2.um_address addr,\n" +
"       (select cz.acct_id, sum(cz.fee) / 100 amount1\n" +
"          from ac2.am_bill_ar_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " cz, pzg1.am_bill_type bity1\n" +
"         where cz.bill_item_id = bity1.bill_item_id\n" +
"           and (bity1.bill_item_name like '%奇艺%' or\n" +
"               bity1.bill_item_name like '%奇异%')\n" +
"         group by cz.acct_id) a1,\n" +
"       (select xz.acct_id, sum(xz.writeoff_fee) / 100 amont2\n" +
"          from (select pxz.acct_id, pxz.bill_item_id, pxz.writeoff_fee\n" +
"                  from ac2.am_writeoff_d pxz\n" +
"                 where pxz.bill_month = " + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "\n" +
"                   and pxz.bill_flag = 'U'\n" +
"                union all\n" +
"                select cxz.acct_id, cxz.bill_item_id, cxz.writeoff_fee\n" +
"                  from ac2.am_writeoff cxz\n" +
"                 where cxz.create_date > date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"                   and cxz.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"                   and cxz.bill_flag = 'U') xz,\n" +
"               pzg1.am_bill_type bity2\n" +
"         where xz.bill_item_id = bity2.bill_item_id\n" +
"           and (bity2.bill_item_name like '%奇艺%' or\n" +
"               bity2.bill_item_name like '%奇异%')\n" +
"         group by xz.acct_id) a2\n" +
" where cust.cust_id = acct.cust_id\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and acct.acct_id = a1.acct_id(+)\n" +
"   and acct.acct_id = a2.acct_id(+)\n" +
"   and addr.expire_date > sysdate\n" +
"   and (a1.amount1 <> 0 or a2.amont2 <> 0)\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name, cust.cust_code";
                int[] columntxt = { 1 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "爱奇艺出销账(去测试客户)--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报电竞出销账(去测试客户)
        public static bool ybaobiao7()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "电竞出销账(去测试客户)--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select  cust.cust_code 客户证号,\n" +
"       part.party_name 姓名,\n" +
"       decode(cust.cust_type,\n" +
"              1,\n" +
"              '公众客户',\n" +
"              2,\n" +
"              '商业客户',\n" +
"              3,\n" +
"              '团体代付客户',\n" +
"              4,\n" +
"              '合同商业客户') 客户类型,\n" +
"       rad.jy_region_name 区域,\n" +
"       addr.std_addr_name 地址,\n" +
"       a1.amount1 出账金额,\n" +
"       a2.amont2 销账金额,\n" +
"                          case\n" +
"                  when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or \n" +
"                        (cust.cust_type = 4 and cust.cust_prop = 10)) then\n" +
"                   '业务用机(免催免停)'\n" +
"                  else\n" +
"                   '普通'\n" +
"                end 客户属性\n" +
"  from cp2.cm_customer cust,\n" +
"       files2.cm_account acct,\n" +
"       cp2.cb_party              part,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       files2.um_address addr,\n" +
"       (select cz.acct_id, sum(cz.fee) / 100 amount1\n" +
"          from ac2.am_bill_ar_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " cz, pzg1.am_bill_type bity1\n" +
"         where cz.bill_item_id = bity1.bill_item_id\n" +
"           and (bity1.bill_item_name like '%电竞%')\n" +
"         group by cz.acct_id) a1,\n" +
"       (select xz.acct_id, sum(xz.writeoff_fee) / 100 amont2\n" +
"          from (select pxz.acct_id, pxz.bill_item_id, pxz.writeoff_fee\n" +
"                  from ac2.am_writeoff_d pxz\n" +
"                 where pxz.bill_month = " + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "\n" +
"                union all\n" +
"                select cxz.acct_id, cxz.bill_item_id, cxz.writeoff_fee\n" +
"                  from ac2.am_writeoff cxz\n" +
"                 where cxz.create_date > date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"                   and cxz.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "') xz,\n" +
"               pzg1.am_bill_type bity2\n" +
"         where xz.bill_item_id = bity2.bill_item_id\n" +
"           and (bity2.bill_item_name like '%电竞%')\n" +
"         group by xz.acct_id) a2\n" +
" where cust.cust_id = acct.cust_id\n" +
"   and cust.cust_id = rad.cust_id\n" +
"   and cust.cust_id = addr.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"   and acct.acct_id = a1.acct_id(+)\n" +
"   and acct.acct_id = a2.acct_id(+)\n" +
"   and (a1.amount1 <> 0 or a2.amont2 <> 0)\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name, cust.cust_code";

                int[] columntxt = { 1 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "电竞出销账(去测试客户)--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报芒果出销账(去测试客户)
        public static bool ybaobiao8()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "芒果出销账(去测试客户)--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select  cust.cust_code 客户证号,\n" +
"       part.party_name 姓名,\n" +
"       decode(cust.cust_type,\n" +
"              1,\n" +
"              '公众客户',\n" +
"              2,\n" +
"              '商业客户',\n" +
"              3,\n" +
"              '团体代付客户',\n" +
"              4,\n" +
"              '合同商业客户') 客户类型,\n" +
"       rad.jy_region_name 区域,\n" +
"       addr.std_addr_name 地址,\n" +
"       a1.amount1 出账金额,\n" +
"       a2.amont2 销账金额,\n" +
"                          case\n" +
"                  when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or \n" +
"                        (cust.cust_type = 4 and cust.cust_prop = 10)) then\n" +
"                   '业务用机(免催免停)'\n" +
"                  else\n" +
"                   '普通'\n" +
"                end 客户属性\n" +
"  from cp2.cm_customer cust,\n" +
"       files2.cm_account acct,\n" +
"       cp2.cb_party              part,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       files2.um_address addr,\n" +
"       (select cz.acct_id, sum(cz.fee) / 100 amount1\n" +
"          from ac2.am_bill_ar_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " cz, pzg1.am_bill_type bity1\n" +
"         where cz.bill_item_id = bity1.bill_item_id\n" +
"           and (bity1.bill_item_name like '%芒果%')\n" +
"         group by cz.acct_id) a1,\n" +
"       (select xz.acct_id, sum(xz.writeoff_fee) / 100 amont2\n" +
"          from (select pxz.acct_id, pxz.bill_item_id, pxz.writeoff_fee\n" +
"                  from ac2.am_writeoff_d pxz\n" +
"                 where pxz.bill_month = " + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "\n" +
"                union all\n" +
"                select cxz.acct_id, cxz.bill_item_id, cxz.writeoff_fee\n" +
"                  from ac2.am_writeoff cxz\n" +
"                 where cxz.create_date > date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"                   and cxz.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "') xz,\n" +
"               pzg1.am_bill_type bity2\n" +
"         where xz.bill_item_id = bity2.bill_item_id\n" +
"           and (bity2.bill_item_name like '%芒果%')\n" +
"         group by xz.acct_id) a2\n" +
" where cust.cust_id = acct.cust_id\n" +
"   and cust.cust_id = rad.cust_id\n" +
"   and cust.cust_id = addr.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"   and acct.acct_id = a1.acct_id(+)\n" +
"   and acct.acct_id = a2.acct_id(+)\n" +
"   and (a1.amount1 <> 0 or a2.amont2 <> 0)\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name, cust.cust_code";

                int[] columntxt = { 1 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "芒果出销账(去测试客户)--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报上文广内容整包出销账(去测试客户)
        public static bool ybaobiao9()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "上文广内容整包出销账(去测试客户)--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select  cust.cust_code 客户证号,\n" +
"       part.party_name 姓名,\n" +
"       decode(cust.cust_type,\n" +
"              1,\n" +
"              '公众客户',\n" +
"              2,\n" +
"              '商业客户',\n" +
"              3,\n" +
"              '团体代付客户',\n" +
"              4,\n" +
"              '合同商业客户') 客户类型,\n" +
"       rad.jy_region_name 区域,\n" +
"       addr.std_addr_name 地址,\n" +
"       a1.amount1 出账金额,\n" +
"       a2.amont2 销账金额,\n" +
"                          case\n" +
"                  when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or \n" +
"                        (cust.cust_type = 4 and cust.cust_prop = 10)) then\n" +
"                   '业务用机(免催免停)'\n" +
"                  else\n" +
"                   '普通'\n" +
"                end 客户属性\n" +
"  from cp2.cm_customer cust,\n" +
"       files2.cm_account acct,\n" +
"       cp2.cb_party              part,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       files2.um_address addr,\n" +
"       (select cz.acct_id, sum(cz.fee) / 100 amount1\n" +
"          from ac2.am_bill_ar_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " cz, pzg1.am_bill_type bity1\n" +
"         where cz.bill_item_id = bity1.bill_item_id\n" +
"           and (bity1.bill_item_name like '%上文广内容整包%')\n" +
"         group by cz.acct_id) a1,\n" +
"       (select xz.acct_id, sum(xz.writeoff_fee) / 100 amont2\n" +
"          from (select pxz.acct_id, pxz.bill_item_id, pxz.writeoff_fee\n" +
"                  from ac2.am_writeoff_d pxz\n" +
"                 where pxz.bill_month = " + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "\n" +
"                union all\n" +
"                select cxz.acct_id, cxz.bill_item_id, cxz.writeoff_fee\n" +
"                  from ac2.am_writeoff cxz\n" +
"                 where cxz.create_date > date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"                   and cxz.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "') xz,\n" +
"               pzg1.am_bill_type bity2\n" +
"         where xz.bill_item_id = bity2.bill_item_id\n" +
"           and (bity2.bill_item_name like '%上文广内容整包%')\n" +
"         group by xz.acct_id) a2\n" +
" where cust.cust_id = acct.cust_id\n" +
"   and cust.cust_id = rad.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id\n" +
"   and acct.acct_id = a1.acct_id(+)\n" +
"   and acct.acct_id = a2.acct_id(+)\n" +
"   and (a1.amount1 <> 0 or a2.amont2 <> 0)\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name, cust.cust_code";

                int[] columntxt = { 1 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "上文广内容整包出销账(去测试客户)--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报炫力出销账(去测试客户)
        public static bool ybaobiao10()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "炫力出销账(去测试客户)--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select  cust.cust_code 客户证号,\n" +
"       part.party_name 姓名,\n" +
"       decode(cust.cust_type,\n" +
"              1,\n" +
"              '公众客户',\n" +
"              2,\n" +
"              '商业客户',\n" +
"              3,\n" +
"              '团体代付客户',\n" +
"              4,\n" +
"              '合同商业客户') 客户类型,\n" +
"       rad.jy_region_name 区域,\n" +
"       addr.std_addr_name 地址,\n" +
"       a1.amount1 出账金额,\n" +
"       a2.amont2 销账金额,\n" +
"                          case\n" +
"                  when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or \n" +
"                        (cust.cust_type = 4 and cust.cust_prop = 10)) then\n" +
"                   '业务用机(免催免停)'\n" +
"                  else\n" +
"                   '普通'\n" +
"                end 客户属性\n" +
"  from cp2.cm_customer cust,\n" +
"       files2.cm_account acct,\n" +
"       cp2.cb_party              part,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       files2.um_address addr,\n" +
"       (select cz.acct_id, sum(cz.fee) / 100 amount1\n" +
"          from ac2.am_bill_ar_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " cz, pzg1.am_bill_type bity1\n" +
"         where cz.bill_item_id = bity1.bill_item_id\n" +
"           and (bity1.bill_item_name like '%炫力%')\n" +
"         group by cz.acct_id) a1,\n" +
"       (select xz.acct_id, sum(xz.writeoff_fee) / 100 amont2\n" +
"          from (select pxz.acct_id, pxz.bill_item_id, pxz.writeoff_fee\n" +
"                  from ac2.am_writeoff_d pxz\n" +
"                 where pxz.bill_month = " + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "\n" +
"                union all\n" +
"                select cxz.acct_id, cxz.bill_item_id, cxz.writeoff_fee\n" +
"                  from ac2.am_writeoff cxz\n" +
"                 where cxz.create_date > date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"                   and cxz.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "') xz,\n" +
"               pzg1.am_bill_type bity2\n" +
"         where xz.bill_item_id = bity2.bill_item_id\n" +
"           and (bity2.bill_item_name like '%炫力%')\n" +
"         group by xz.acct_id) a2\n" +
" where cust.cust_id = acct.cust_id\n" +
"   and cust.cust_id = rad.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id\n" +
"   and acct.acct_id = a1.acct_id(+)\n" +
"   and acct.acct_id = a2.acct_id(+)\n" +
"   and (a1.amount1 <> 0 or a2.amont2 <> 0)\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name, cust.cust_code";

                int[] columntxt = { 1 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "炫力出销账(去测试客户)--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报学霸宝盒出销账(去测试客户)
        public static bool ybaobiao11()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "学霸宝盒出销账(去测试客户)--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select  cust.cust_code 客户证号,\n" +
"       part.party_name 姓名,\n" +
"       decode(cust.cust_type,\n" +
"              1,\n" +
"              '公众客户',\n" +
"              2,\n" +
"              '商业客户',\n" +
"              3,\n" +
"              '团体代付客户',\n" +
"              4,\n" +
"              '合同商业客户') 客户类型,\n" +
"       rad.jy_region_name 区域,\n" +
"       addr.std_addr_name 地址,\n" +
"       a1.amount1 出账金额,\n" +
"       a2.amont2 销账金额,\n" +
"                          case\n" +
"                  when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or \n" +
"                        (cust.cust_type = 4 and cust.cust_prop = 10)) then\n" +
"                   '业务用机(免催免停)'\n" +
"                  else\n" +
"                   '普通'\n" +
"                end 客户属性\n" +
"  from cp2.cm_customer cust,\n" +
"       files2.cm_account acct,\n" +
"       cp2.cb_party              part,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       files2.um_address addr,\n" +
"       (select cz.acct_id, sum(cz.fee) / 100 amount1\n" +
"          from ac2.am_bill_ar_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " cz, pzg1.am_bill_type bity1\n" +
"         where cz.bill_item_id = bity1.bill_item_id\n" +
"           and (bity1.bill_item_name like '%学霸宝盒%')\n" +
"         group by cz.acct_id) a1,\n" +
"       (select xz.acct_id, sum(xz.writeoff_fee) / 100 amont2\n" +
"          from (select pxz.acct_id, pxz.bill_item_id, pxz.writeoff_fee\n" +
"                  from ac2.am_writeoff_d pxz\n" +
"                 where pxz.bill_month = " + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "\n" +
"                union all\n" +
"                select cxz.acct_id, cxz.bill_item_id, cxz.writeoff_fee\n" +
"                  from ac2.am_writeoff cxz\n" +
"                 where cxz.create_date > date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"                   and cxz.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "') xz,\n" +
"               pzg1.am_bill_type bity2\n" +
"         where xz.bill_item_id = bity2.bill_item_id\n" +
"           and (bity2.bill_item_name like '%学霸宝盒%')\n" +
"         group by xz.acct_id) a2\n" +
" where cust.cust_id = acct.cust_id\n" +
"   and cust.cust_id = rad.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id\n" +
"   and acct.acct_id = a1.acct_id(+)\n" +
"   and acct.acct_id = a2.acct_id(+)\n" +
"   and (a1.amount1 <> 0 or a2.amont2 <> 0)\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name, cust.cust_code";

                int[] columntxt = { 1 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "学霸宝盒出销账(去测试客户)--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报江阴互动基本缴费用户数及终端数
        public static bool ybaobiao12()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "江阴互动基本缴费用户数及终端数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select rad.jy_region_name 区域,\n" +
"       decode(cust.cust_type,\n" +
"              1,\n" +
"              '公众客户',\n" +
"              2,\n" +
"              '商业客户',\n" +
"              3,\n" +
"              '团体代付客户',\n" +
"              4,\n" +
"              '合同商业客户') 客户类型,\n" +
"       count(distinct cust.cust_code) 江阴互动基本缴费客户数,\n" +
"       count(distinct term.serial_no) 江阴互动基本缴费机顶盒数\n" +
"  from rep2.rep_fact_cm_customer_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "   cust,\n" +
"       rep2.rep_fact_um_subscriber_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " subs,\n" +
"       res1.res_terminal                    term,\n" +
"       res1.res_sku                         rsku,\n" +
"       wxjy.jy_region_address_rel           rad\n" +
" where cust.cust_id = subs.cust_id\n" +
"   and subs.bill_id = term.serial_no\n" +
"   and term.res_sku_id = rsku.res_sku_id\n" +
"   and cust.cust_id = rad.cust_id\n" +
"   and rsku.res_type_id = 2\n" +
"   and subs.is_dtv = 1\n" +
"   and subs.is_dbitv = 1 --互动\n" +
"   and subs.is_dbitv_paied = 1 --互动缴费\n" +
"   and cust.corp_org_id = 3328\n" +
"   and rsku.res_sku_name in ('银河高清基本型HDC6910(江阴)',\n" +
"                             '银河高清交互型HDC691033(江阴)',\n" +
"                             '银河智能高清交互型HDC6910798(江阴)',\n" +
"                             '银河4K交互型HDC691090',\n" +
"                             '4K超高清型II型融合型（EOC）',\n" +
"                             '银河4K超高清II型融合型HDC6910B1(江阴)',\n" +
"                             '4K超高清简易型（基本型）')\n" +
"   and exists\n" +
" (select 1\n" +
"          from rep2.rep_fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " insp,\n" +
"               wxjy.jy_dbitv_product_local       hd\n" +
"         where insp.srvpkg_id = hd.offer_id\n" +
"           and insp.subscriber_ins_id = subs.subscriber_ins_id)\n" +
" group by rad.jy_region_name,\n" +
"          decode(cust.cust_type,\n" +
"                 1,\n" +
"                 '公众客户',\n" +
"                 2,\n" +
"                 '商业客户',\n" +
"                 3,\n" +
"                 '团体代付客户',\n" +
"                 4,\n" +
"                 '合同商业客户')\n" +
" order by rad.jy_region_name,\n" +
"          decode(cust.cust_type,\n" +
"                 1,\n" +
"                 '公众客户',\n" +
"                 2,\n" +
"                 '商业客户',\n" +
"                 3,\n" +
"                 '团体代付客户',\n" +
"                 4,\n" +
"                 '合同商业客户')";

                int[] columntxt = { 0 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "江阴互动基本缴费用户数及终端数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报集团和商业全量终端明细
        public static bool ybaobiao13()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "集团和商业全量终端明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                term.serial_no 机顶盒,\n" +
"                subs.sub_bill_id 智能卡,\n" +
"                rsku.res_sku_name 资源类型,\n" +
"                ures.create_date 资源订购时间,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址\n" +
"  from cp2.cm_customer            cust,\n" +
"       files2.cm_account          acct,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_res              ures,\n" +
"       res1.res_terminal          term,\n" +
"       res1.res_sku               rsku,\n" +
"       files2.um_address          addr,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       cp2.cb_party              part,\n" +
"       wxjy.jy_contact_rel        rel\n" +
" where cust.cust_id = acct.cust_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"   and subs.subscriber_ins_id = ures.subscriber_ins_id\n" +
"   and ures.res_equ_no = term.serial_no\n" +
"   and term.res_sku_id = rsku.res_sku_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and addr.expire_date > sysdate\n" +
"   and ures.expire_date > sysdate\n" +
"   and rsku.res_type_id = 2\n" +
"   and cust.cust_type <> 1\n" +
"   and cust.own_corp_org_id = 3328\n" +
"   and exists\n" +
" (select 1\n" +
"          from files2.um_offer_06 ofer, files2.um_offer_sta_02 fsta\n" +
"         where ofer.offer_ins_id = fsta.offer_ins_id\n" +
"           and ofer.prod_service_id = 1002\n" +
"           and fsta.offer_status = '1'\n" +
"           and fsta.os_status is null\n" +
"           and ofer.subscriber_ins_id = subs.subscriber_ins_id)\n" +
" order by cust.cust_code, term.serial_no";

                int[] columntxt = { 1, 4, 5 }; //哪些列是文本格式
                int[] columndate = { 7, 8 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "集团和商业全量终端明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报当月宽带新增用户明细
        public static bool ybaobiao15()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "当月宽带新增用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                subs.login_name 宽带登录名,\n" +
"                subs.create_date 开户时间,\n" +
"                updf.offer_name 产品名称,\n" +
"                fsta.create_date 产品订购时间,\n" +
"                rel.cont_number || ',' || rel.cont_number2 联系方式,\n" +
"                rel.family_number || ',' || rel.family_number2 联系方式2,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址\n" +
"  from cp2.cm_customer            cust,\n" +
"       files2.cm_account          acct,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_offer_06         ofer,\n" +
"       files2.um_offer_sta_02     fsta,\n" +
"       upc1.pm_offer              updf,\n" +
"       files2.um_address          addr,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       cp2.cb_party              part,\n" +
"       wxjy.jy_contact_rel        rel\n" +
" where cust.cust_id = subs.cust_id\n" +
"   and cust.cust_id = acct.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"   and fsta.offer_id = updf.offer_id\n" +
"   and addr.expire_date > sysdate\n" +
"   and subs.login_name is not null\n" +
"   and subs.main_spec_id <> 80020199 --去除虚用户\n" +
"   and ofer.prod_service_id = 1004\n" +
"   and fsta.offer_status = '1'\n" +
"   and fsta.os_status is null\n" +
"   and subs.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"   and subs.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by cust.cust_code, rad.jy_region_name";

                int[] columntxt = { 1, 4, 9 }; //哪些列是文本格式
                int[] columndate = { 5, 7 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "当月宽带新增用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion
 
        #region 月报今年全量复通加新开户用户明细
        public static bool ybaobiao17()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "今年全量复通加新开户用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct cust.cust_code      客户证号,\n" +
"                cust.create_date    开户时间,\n" +
"                cust.cust_name      姓名,\n" +
"                rad.jy_region_name  区域,\n" +
"                cust.cust_cert_addr 地址,\n" +
"                info.cont_number    移动号码1,\n" +
"                info.family_number  家庭电话1,\n" +
"                info.cont_number2   移动号码2,\n" +
"                info.family_number2 家庭电话2,\n" +
"                 wgrel.grid_name 网格\n" +
"  from rep2.rep_fact_cm_customer_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "   cust,\n" +
"       rep2.rep_fact_um_subscriber_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " subs,\n" +
"       rep2.rep_fact_cust_info_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "     info,\n" +
"       wxjy.jy_region_address_rel           rad,\n" +
"           wxjy.jy_customer_wg_rel wgrel\n" +
" where cust.cust_id = rad.cust_id(+)\n" +
"   and cust.cust_id = info.cust_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and cust.cust_id = wgrel.cust_id(+)\n" +
"   and subs.is_dtv = 1\n" +
"   and subs.is_paied = 1\n" +
"   and cust.corp_org_id = 3328\n" +
"   and cust.cust_code not in\n" +
"       (select distinct a.cust_code\n" +
"          from rep2.rep_fact_cm_customer_20221231       a,\n" +
"                rep2.rep_fact_um_subscriber_20221231   b\n" +
"         where a.cust_id = b.cust_id\n" +
"           and b.is_dtv = 1\n" +
"           and b.is_paied = 1\n" +
"           and a.corp_org_id = 3328)\n" +
" order by rad.jy_region_name, cust.cust_code";

                int[] columntxt = { 1, 6, 7, 8, 9 }; //哪些列是文本格式
                int[] columndate = { 2 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "今年全量复通加新开户用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报今年全量流失用户明细
        public static bool ybaobiao18()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "今年全量流失用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                term.serial_no 机顶盒,\n" +
"                subs.sub_bill_id 智能卡,\n" +
"                case\n" +
"                  when nvl(subs.main_subscriber_ins_id, 0) = 0 then\n" +
"                   '主机'\n" +
"                  else\n" +
"                   ''\n" +
"                end 主副机,\n" +
"                case\n" +
"                  when subs.state = 'M' then\n" +
"                   '暂停'\n" +
"                  when fsta.os_status = '1' then\n" +
"                   '欠费停机'\n" +
"                  when fsta.os_status in ('3', '4') then\n" +
"                   '暂停'\n" +
"                  else\n" +
"                   ''\n" +
"                end 停开机状态,\n" +
"                case\n" +
"                  when subs.state = 'M' then\n" +
"                   subs.done_date\n" +
"                  when fsta.os_status = '1' then\n" +
"                   fsta.done_date\n" +
"                  when fsta.os_status in (3, 4) then\n" +
"                   subs.done_date\n" +
"                  else\n" +
"                   null\n" +
"                end 受理时间,\n" +
"                rsku.res_sku_name 资源类型,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_offer_06         ofer,\n" +
"       files2.um_offer_sta_02     fsta,\n" +
"       files2.um_res              ures,\n" +
"       res1.res_terminal          term,\n" +
"       res1.res_sku               rsku,\n" +
"       files2.um_address          addr,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_customer_wg_rel    wg\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ures.subscriber_ins_id\n" +
"   and ures.res_equ_no = term.serial_no\n" +
"   and term.res_sku_id = rsku.res_sku_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"   and nvl(subs.main_subscriber_ins_id, 0) = 0\n" +
"   and ofer.prod_service_id = 1002 --- 基本节目\n" +
"   and subs.main_spec_id = 80020001 ---数字电视\n" +
"   and addr.expire_date > sysdate\n" +
"   and ures.expire_date > sysdate\n" +
"   and fsta.expire_date > sysdate\n" +
"   and ofer.expire_date > sysdate\n" +
"   and rsku.res_type_id = 2\n" +
"   and cust.cust_type = 1 and cust.cust_prop <> '6'\n" +
"   and cust.own_corp_org_id = 3328\n" +
"   and not exists\n" +
" (select 1\n" +
"          from files2.um_offer_06 ofer1, files2.um_offer_sta_02 fsta1\n" +
"         where ofer1.offer_ins_id = fsta1.offer_ins_id\n" +
"           and ofer1.prod_service_id = 1002\n" +
"           and fsta1.offer_status = '1'\n" +
"           and fsta1.os_status is null\n" +
"           and fsta1.expire_date > sysdate\n" +
"           and ofer1.expire_date > sysdate\n" +
"           and ofer1.subscriber_ins_id = subs.subscriber_ins_id)\n" +
"   and exists (select 1\n" +
"          from wxjy.jy_cust_tmp_20211231 a\n" +
"         where a.cust_code = cust.cust_code)\n" +
" order by rad.jy_region_name, cust.cust_code, term.serial_no";
                int[] columntxt = { 1, 4, 5, 10 }; //哪些列是文本格式
                int[] columndate = { 8 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "今年全量流失用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报有线网络总客户数
        public static bool ybaobiao19()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "有线网络总客户数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select count(distinct a.cust_code) 有线网络总客户数\n" +
"  from cp2.cm_customer a\n" +
" where a.cust_code is not null\n" +
"   and a.own_corp_org_id = 3328";
                int[] columntxt = { 0 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "有线网络总客户数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报全量客户明细
        public static bool ybaobiao20()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "全量客户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct  cust.cust_code  客户证号, cust.cust_name 姓名 ,\n" +
"           rad.jy_region_name 区域 ,rad.std_addr_name 地址 ,\n" +
"            decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                         case when cust.cust_type = 1 and cust.cust_prop = 6 then '测试客户' else '' end  只判断公众是否测试\n" +
"  from rep2.rep_fact_cm_customer_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "   cust,\n" +
"          files2.cm_account          acct,\n" +
"       files2.um_address       addr, wxjy.jy_region_address_rel rad\n" +
" where  cust.cust_id=acct.cust_id\n" +
"   and cust.cust_id = addr.cust_id\n" +
"   and  cust.corp_org_id = 3328  and cust.cust_code is not null\n" +
"   and cust.cust_id=rad.cust_id\n" +
"   and addr.expire_date >= sysdate";
                int[] columntxt = { 1 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "全量客户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报爱奇艺销账金额明细数据（加科目且去测试客户）
        public static bool ybaobiao21()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "爱奇艺销账金额明细数据（加科目且去测试客户）--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select cust.cust_code 客户证号,\n" +
"       part.party_name 姓名,\n" +
"       decode(cust.cust_type,\n" +
"              1,\n" +
"              '公众客户',\n" +
"              2,\n" +
"              '商业客户',\n" +
"              3,\n" +
"              '团体代付客户',\n" +
"              4,\n" +
"              '合同商业客户') 客户类型,\n" +
"       rad.jy_region_name 区域,\n" +
"       addr.std_addr_name 地址,\n" +
"       bilt.bill_item_name 账单科目,\n" +
"       sifo.service_name 业务,\n" +
"       sum(xz.writeoff_fee) / 100 销账金额,\n" +
"                          case\n" +
"                  when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or \n" +
"                        (cust.cust_type = 4 and cust.cust_prop = 10)) then\n" +
"                   '业务用机(免催免停)'\n" +
"                  else\n" +
"                   '普通'\n" +
"                end 客户属性\n" +
"  from cp2.cm_customer            cust,\n" +
"       files2.cm_account          acct,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       cp2.cb_party              part,\n" +
"       files2.um_address          addr,\n" +
"       (select pxz.acct_id,\n" +
"               pxz.bill_item_id,\n" +
"               pxz.asset_item_id,\n" +
"               pxz.writeoff_fee\n" +
"          from ac2.am_writeoff_d pxz\n" +
"         where pxz.bill_month = " + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "\n" +
"           and pxz.bill_flag = 'U'\n" +
"        union all\n" +
"        select cxz.acct_id,\n" +
"               cxz.bill_item_id,\n" +
"               cxz.asset_item_id,\n" +
"               cxz.writeoff_fee\n" +
"          from ac2.am_writeoff cxz\n" +
"         where cxz.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"           and cxz.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"           and cxz.bill_flag = 'U') xz,\n" +
"       pzg1.am_bill_type    bilt,\n" +
"       pzg1.am_item_service aits,\n" +
"       pzg1.am_service_info sifo\n" +
" where cust.cust_id = acct.cust_id\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.party_id = part.party_id\n" +
"   and acct.acct_id = xz.acct_id\n" +
"   and xz.bill_item_id = bilt.bill_item_id\n" +
"   and bilt.bill_item_id = aits.am_item_type_id\n" +
"   and aits.service_id = sifo.service_id\n" +
"   and xz.writeoff_fee <> 0\n" +
"   and addr.expire_date > sysdate\n" +
"   and (bilt.bill_item_name like '%奇异%' or bilt.bill_item_name like '%奇艺%')\n" +
"   and cust.own_corp_org_id = 3328\n" +
" group by cust.cust_code,\n" +
"          part.party_name,\n" +
"          decode(cust.cust_type,\n" +
"                 1,\n" +
"                 '公众客户',\n" +
"                 2,\n" +
"                 '商业客户',\n" +
"                 3,\n" +
"                 '团体代付客户',\n" +
"                 4,\n" +
"                 '合同商业客户'),\n" +
"          rad.jy_region_name,\n" +
"          addr.std_addr_name,\n" +
"          bilt.bill_item_name,\n" +
"          sifo.service_name,\n" +
"                          case\n" +
"                  when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or \n" +
"                       (cust.cust_type = 4 and cust.cust_prop = 10)) then\n" +
"                   '业务用机(免催免停)'\n" +
"                  else\n" +
"                   '普通'\n" +
"                end \n" +
" order by rad.jy_region_name, cust.cust_code";

                int[] columntxt = { 1 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "爱奇艺销账金额明细数据（加科目且去测试客户）--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报360元分2和5年及720元分3和5年返还及押金返点的客户返充明细数据
        public static bool ybaobiao22()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "360元分2和5年及720元分3和5年返还及押金返点的客户返充明细数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                to_char(pay.business_id) 流水号,\n" +
"                pay.amount / 100 金额,\n" +
"                pay.create_date 返还时间,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址\n" +
"  from cp2.cm_customer            cust,\n" +
"       files2.cm_account          acct,\n" +
"       files2.um_address          addr,\n" +
"       ac2.am_busi                pay,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       cp2.cb_party              part,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where cust.cust_id = acct.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and acct.acct_id = pay.acct_id\n" +
"   and pay.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"   and pay.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"   and pay.asset_item_id = 100\n" +
"   and pay.amount <> 0\n" +
"   and pay.business_type_id = 4302\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by cust.cust_code";

                int[] columntxt = { 1, 4, 7 }; //哪些列是文本格式
                int[] columndate = { 6 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "360元分2和5年及720元分3和5年返还及押金返点的客户返充明细数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报现金续费返充明细（360、720、返点）
        public static bool ybaobiao23()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "现金续费返充明细（360、720、返点）--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct prod.offer_name 产品,\n" +
"                aapl.origin_balance / 100 返还总额_元,\n" +
"                aapl.months 返还总次数,\n" +
"                aapl.start_month 开始返还账期,\n" +
"                aapl.remain_months 剩余次数,\n" +
"                aapl.month_return_amount / 100 每次返还金额_元,\n" +
"                aapl.remain_balance / 100 剩余金额_元,\n" +
"                cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址\n" +
"  from cp2.cm_customer            cust,\n" +
"       files2.cm_account          acct,\n" +
"       files2.um_address          addr,\n" +
"       files2.um_subscriber       subs,\n" +
"       ac2.am_apportion           appo,\n" +
"       ac2.am_apportion_plan      aapl,\n" +
"       upc1.pm_offer              prod,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       cp2.cb_party              part,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where cust.cust_id = acct.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = appo.subscriber_id\n" +
"   and appo.price_ins_id = aapl.price_ins_id\n" +
"   and aapl.offer_id = prod.offer_id\n" +
"   and addr.expire_date > sysdate\n" +
"   and aapl.month_return_amount > 0\n" +
"   and aapl.state = 'R'\n" +
"   and aapl.remain_months = 1\n" +
"   and cust.own_corp_org_id = 3328\n" +
"   and appo.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"   and appo.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"   and (prod.offer_name like '%预存360元分2年返还套餐%'\n" +
"   or prod.offer_name like '%高清互动升级充值360元分5年返还套餐%'\n" +
"   or prod.offer_name like '%720元3年促销%'\n" +
"   or prod.offer_name like '%720元5年促销%'\n" +
"   or prod.offer_name like '%宽带猫押金返点%'\n" +
"   or prod.offer_name like '%设备押金返点%')\n" +
" order by prod.offer_name, cust.cust_code";

                int[] columntxt = { 8, 11 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "现金续费返充明细（360、720、返点）--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报爱奇艺出账金额明细数据（加科目且去测试客户）
        public static bool ybaobiao24()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "爱奇艺出账金额明细数据（加科目且去测试客户）--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                a1.bill_item_name 出账科目,\n" +
"                a1.amount1 出账金额,\n" +
"                          case\n" +
"                  when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or \n" +
"                        (cust.cust_type = 4 and cust.cust_prop = 10)) then\n" +
"                   '业务用机(免催免停)'\n" +
"                  else\n" +
"                   '普通'\n" +
"                end 客户属性\n" +
"  from cp2.cm_customer            cust,\n" +
"       files2.cm_account          acct,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       cp2.cb_party              part,\n" +
"       files2.um_address          addr,\n" +
"       (select cz.acct_id, bity1.bill_item_name, sum(cz.fee) / 100 amount1\n" +
"          from ac2.am_bill_ar_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " cz, pzg1.am_bill_type bity1\n" +
"         where cz.bill_item_id = bity1.bill_item_id\n" +
"           and (bity1.bill_item_name like '%奇异%' or\n" +
"               bity1.bill_item_name like '%奇艺%')\n" +
"         group by cz.acct_id, bity1.bill_item_name) a1\n" +
" where cust.cust_id = acct.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and acct.acct_id = a1.acct_id(+)\n" +
"   and a1.amount1 <> 0\n" +
"   and addr.expire_date > sysdate\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name, cust.cust_code";

                int[] columntxt = { 1 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "爱奇艺出账金额明细数据（加科目且去测试客户）--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报纯标清缴费用户明细带测试
        public static bool ybaobiao25()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "纯标清缴费用户明细带测试--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                decode(rela.state, 0, '是', 1, '') 是否银行代扣,\n" +
"                      case when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or (cust.cust_type = 4 and cust.cust_prop = 10)) then '测试客户' else '' end  只判断公众是否测试,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格\n" +
"  from cp2.cm_customer            cust,\n" +
"       files2.cm_account          acct,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_res              ures,\n" +
"       res1.res_terminal          term,\n" +
"       res1.res_sku               rsku,\n" +
"       files2.um_offer_06         ofer,\n" +
"       files2.um_offer_sta_02     fsta,\n" +
"       files2.um_address          addr,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_customer_wg_rel    wg,\n" +
"       ac2.am_entrust_relation    rela\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = acct.cust_id\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and acct.acct_id = rela.acct_id(+)\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"   and subs.subscriber_ins_id = ures.subscriber_ins_id\n" +
"   and ures.res_equ_no = term.serial_no\n" +
"   and term.res_sku_id = rsku.res_sku_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and rsku.res_type_id = 2\n" +
"   and ofer.prod_service_id = 1002\n" +
"   and ures.expire_date > sysdate\n" +
"   and addr.expire_date > sysdate\n" +
"   and fsta.expire_date > sysdate\n" +
"   and ofer.expire_date > sysdate\n" +
"   and cust.own_corp_org_id = 3328\n" +
"   and not exists\n" +
" (select 1\n" +
"          from files2.um_subscriber subs1,\n" +
"               files2.um_res        ures1,\n" +
"               res1.res_terminal    term1,\n" +
"               res1.res_sku         rsku1\n" +
"         where subs1.subscriber_ins_id = ures1.subscriber_ins_id\n" +
"           and ures1.res_equ_no = term1.serial_no\n" +
"           and term1.res_sku_id = rsku1.res_sku_id\n" +
"           and ures1.res_type_id = 2\n" +
"           and rsku1.res_sku_name in ('银河高清基本型HDC6910(江阴)',\n" +
"                                      '银河高清交互型HDC691033(江阴)',\n" +
"                                      '银河智能高清交互型HDC6910798(江阴)',\n" +
"                                      '银河4K交互型HDC691090',\n" +
"                                      '4K超高清型II型融合型（EOC）',\n" +
"                                      '银河4K超高清II型融合型HDC6910B1(江阴)',\n" +
"                                      '4K超高清简易型（基本型）')\n" +
"           and subs1.cust_id = cust.cust_id)\n" +
"   and nvl(subs.main_subscriber_ins_id, 0) = 0 ---主机\n" +
"   and exists\n" +
" (select 1\n" +
"          from files2.um_offer_06 ofer2, files2.um_offer_sta_02 fsta2\n" +
"         where ofer2.offer_ins_id = fsta2.offer_ins_id\n" +
"           and fsta2.expire_date > sysdate\n" +
"           and ofer2.expire_date > sysdate\n" +
"           and fsta2.offer_status = '1'\n" +
"           and ofer2.prod_service_id = 1002\n" +
"           and fsta2.os_status is null\n" +
"           and ofer2.subscriber_ins_id = subs.subscriber_ins_id) --基本包为开通";
                int[] columntxt = { 1, 5, 6, 10 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "纯标清缴费用户明细带测试--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报电子渠道缴费明细
        public static bool ybaobiao26()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "电子渠道缴费明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                to_char(pay.business_id) 缴费编号,\n" +
"                pay.amount / 100 金额,\n" +
"                pay.create_date 时间,\n" +
"                dtl.bank_name 渠道,\n" +
"                pay.peer_business_id 渠道编号,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格\n" +
"  from cp2.cm_customer            cust,\n" +
"       files2.cm_account          acct,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_address          addr,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_customer_wg_rel    wg,\n" +
"       ac2.am_busi                pay,\n" +
"       ac2.am_busi_ext            dtl,\n" +
"       pzg1.bs_channel            cha\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = acct.cust_id\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and acct.acct_id = pay.acct_id\n" +
"   and pay.business_id = dtl.business_id\n" +
"   and pay.channel_id = cha.channel_id\n" +
"   and addr.expire_date > sysdate\n" +
"   and pay.channel_id <> 0\n" +
"   and pay.peer_business_id is not null\n" +
"   and pay.cancel_flag = 'U'\n" +
"   and pay.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"   and pay.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name, cust.cust_code, pay.create_date";

                int[] columntxt = { 1, 4, 8, 9 }; //哪些列是文本格式
                int[] columndate = { 6 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "电子渠道缴费明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报果果乐园探奇动物界出销账金额明细数据(去测试客户)
        public static bool ybaobiao28()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "果果乐园探奇动物界出销账金额明细数据(去测试客户)--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                a1.amount1 出账金额,\n" +
"                a2.amont2 销账金额,\n" +
"                          case\n" +
"                  when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or \n" +
"                        (cust.cust_type = 4 and cust.cust_prop = 10)) then\n" +
"                   '业务用机(免催免停)'\n" +
"                  else\n" +
"                   '普通'\n" +
"                end 客户属性\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.cm_account          acct,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       files2.um_address          addr,\n" +
"       (select cz.acct_id, sum(cz.fee) / 100 amount1\n" +
"          from ac2.am_bill_ar_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " cz, pzg1.am_bill_type bity1\n" +
"         where cz.bill_item_id = bity1.bill_item_id\n" +
"           and (bity1.bill_item_name like '%果果乐园%' or bity1.bill_item_name like '%探奇动物界%')\n" +
"         group by cz.acct_id) a1,\n" +
"       (select xz.acct_id, sum(xz.writeoff_fee) / 100 amont2\n" +
"          from (select pxz.acct_id, pxz.bill_item_id, pxz.writeoff_fee\n" +
"                  from ac2.am_writeoff_d pxz\n" +
"                 where pxz.bill_month = " + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "\n" +
"                   and pxz.bill_flag = 'U'\n" +
"                union all\n" +
"                select cxz.acct_id, cxz.bill_item_id, cxz.writeoff_fee\n" +
"                  from ac2.am_writeoff cxz\n" +
"                 where cxz.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"                   and cxz.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"                   and cxz.bill_flag = 'U') xz,\n" +
"               pzg1.am_bill_type bity2\n" +
"         where xz.bill_item_id = bity2.bill_item_id\n" +
"           and (bity2.bill_item_name like '%果果乐园%' or bity2.bill_item_name like '%探奇动物界%')\n" +
"         group by xz.acct_id) a2\n" +
" where cust.cust_id = acct.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and acct.acct_id = a1.acct_id(+)\n" +
"   and acct.acct_id = a2.acct_id(+)\n" +
"   and (a1.amount1 <> 0 or a2.amont2 <> 0)\n" +
"   and addr.expire_date > sysdate\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name, cust.cust_code";

                int[] columntxt = { 1 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "果果乐园探奇动物界出销账金额明细数据(去测试客户)--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报极视影院出销账金额明细数据(去测试客户)
        public static bool ybaobiao29()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "极视影院出销账金额明细数据(去测试客户)--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                a1.amount1 出账金额,\n" +
"                a2.amont2 销账金额,\n" +
"                          case\n" +
"                  when ((cust.cust_type = 1 and cust.cust_prop = '6') or  (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or \n" +
"                        (cust.cust_type = 4 and cust.cust_prop = 10)) then\n" +
"                   '业务用机(免催免停)'\n" +
"                  else\n" +
"                   '普通'\n" +
"                end 客户属性\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.cm_account          acct,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       files2.um_address          addr,\n" +
"       (select cz.acct_id, sum(cz.fee) / 100 amount1\n" +
"          from ac2.am_bill_ar_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " cz, pzg1.am_bill_type bity1\n" +
"         where cz.bill_item_id = bity1.bill_item_id\n" +
"           and (bity1.bill_item_name like '%极视影院%')\n" +
"         group by cz.acct_id) a1,\n" +
"       (select xz.acct_id, sum(xz.writeoff_fee) / 100 amont2\n" +
"          from (select pxz.acct_id, pxz.bill_item_id, pxz.writeoff_fee\n" +
"                  from ac2.am_writeoff_d pxz\n" +
"                 where pxz.bill_month = " + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "\n" +
"                   and pxz.bill_flag = 'U'\n" +
"                union all\n" +
"                select cxz.acct_id, cxz.bill_item_id, cxz.writeoff_fee\n" +
"                  from ac2.am_writeoff cxz\n" +
"                 where cxz.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"                   and cxz.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"                   and cxz.bill_flag = 'U') xz,\n" +
"               pzg1.am_bill_type bity2\n" +
"         where xz.bill_item_id = bity2.bill_item_id\n" +
"           and (bity2.bill_item_name like '%极视影院%')\n" +
"         group by xz.acct_id) a2\n" +
" where cust.cust_id = acct.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and acct.acct_id = a1.acct_id(+)\n" +
"   and acct.acct_id = a2.acct_id(+)\n" +
"   and (a1.amount1 <> 0 or a2.amont2 <> 0)\n" +
"   and addr.expire_date > sysdate\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name, cust.cust_code";

                int[] columntxt = { 1 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "极视影院出销账金额明细数据(去测试客户)--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报现金通用账本余额不足30元的客户明细数据
        public static bool ybaobiao30()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "现金通用账本余额不足30元的客户明细数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct rad.jy_region_name 区域,\n" +
"                cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                tmp.asset_item_name 账本,\n" +
"                tmp.amount 余额,\n" +
"                case\n" +
"                  when exists (select 1\n" +
"                          from files2.um_subscriber subs1\n" +
"                         where subs1.main_spec_id = 80020003\n" +
"                           and subs1.login_name is not null\n" +
"                           and subs1.cust_id = cust.cust_id) then\n" +
"                   '是'\n" +
"                  else\n" +
"                   ''\n" +
"                end 是否有宽带,\n" +
"                case\n" +
"                  when exists (select 1\n" +
"                          from files2.um_subscriber subs2,\n" +
"                               files2.um_res        ures2,\n" +
"                               res2.res_terminal    term2,\n" +
"                               res2.res_sku         rsku2\n" +
"                         where subs2.subscriber_ins_id =\n" +
"                               ures2.subscriber_ins_id\n" +
"                           and ures2.res_equ_no = term2.serial_no\n" +
"                           and term2.res_sku_id = rsku2.res_sku_id\n" +
"                           and rsku2.res_type_id = 2\n" +
"                           and ures2.expire_date > sysdate\n" +
"                           and rsku2.res_sku_name in\n" +
"                               ('银河高清基本型HDC6910(江阴)',\n" +
"                                '银河高清交互型HDC691033(江阴)',\n" +
"                                '银河智能高清交互型HDC6910798(江阴)',\n" +
"                                '银河4K交互型HDC691090',\n" +
"                                '4K超高清型II型融合型（EOC）',\n" +
"                                '银河4K超高清II型融合型HDC6910B1(江阴)',\n" +
"                                '4K超高清简易型（基本型）')\n" +
"                           and subs2.cust_id = cust.cust_id) then\n" +
"                   '有'\n" +
"                  else\n" +
"                   ''\n" +
"                end 是否存在高清机顶盒\n" +
"  from cp2.cm_customer cust,\n" +
"        cp2.cb_party part,\n" +
"       (select acct.cust_id,\n" +
"               acct.acct_id,\n" +
"               acct.acct_name,\n" +
"               nvl(aaty.asset_item_id, 100) asset_item_id,\n" +
"               nvl(aaty.asset_item_name, '通用现金账本') asset_item_name,\n" +
"               nvl(sum(blac.balance), 0) / 100 amount\n" +
"          from files2.cm_account  acct,\n" +
"               ac2.am_balance_" + DateTime.Now.ToString("MM") + "  blac, --每月一张表\n" +
"               pzg1.am_asset_type aaty\n" +
"         where acct.acct_id = blac.acct_id(+)\n" +
"           and blac.asset_item_id = aaty.asset_item_id(+)\n" +
"           and acct.acct_name not like '%测试%'\n" +
"           and acct.acct_name not like '%ceshi%'\n" +
"           and acct.acct_name not like '%test%'\n" +
"           and acct.corp_org_id = 3328\n" +
"         group by acct.cust_id,\n" +
"                  acct.acct_id,\n" +
"                  acct.acct_name,\n" +
"                  nvl(aaty.asset_item_id, 100),\n" +
"                  nvl(aaty.asset_item_name, '通用现金账本')) tmp,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       files2.um_address addr,\n" +
"       wxjy.jy_contact_rel rel\n" +
" where cust.cust_id = tmp.cust_id\n" +
"   and cust.party_id = part.party_id\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_type = 1 --只统计公众客户\n" +
"   and cust.cust_prop <> '6' --去除免费客户\n" +
"   and tmp.asset_item_id = 100 --通用现金账本\n" +
"   and addr.expire_date > sysdate\n" +
"   and tmp.amount >= 0\n" +
"   and tmp.amount < 30\n" +
"   and cust.own_corp_org_id = 3328\n" +
"   and not exists (select 1\n" +
"          from ac2.am_entrust_relation rela\n" +
"         where rela.exp_date > sysdate\n" +
"           and rela.state = 0\n" +
"           and rela.acct_id = tmp.acct_id) --不是银行托收客户\n" +
"   and exists (select 1\n" +
"          from files2.um_subscriber   subs,\n" +
"               files2.um_offer_06     ofer,\n" +
"               files2.um_offer_sta_02 fsta\n" +
"         where subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"           and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"           and subs.main_spec_id = 80020001\n" +
"           and ofer.expire_date > sysdate\n" +
"           and fsta.expire_date > sysdate\n" +
"           and ofer.prod_service_id = 1002\n" +
"           and fsta.offer_status = '1'\n" +
"           and fsta.os_status is null\n" +
"           and subs.cust_id = cust.cust_id) --有开通的缴费机顶盒\n" +
" order by rad.jy_region_name, addr.std_addr_name, cust.cust_code";

                int[] columntxt = { 2, 5 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "现金通用账本余额不足30元的客户明细数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报营业员异地受理明细数据
        public static bool ybaobiao31()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "营业员异地受理明细数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select tmp.*\n" +
"  from (select distinct cust.cust_code 客户证号,\n" +
"                        part.party_name 姓名,\n" +
"                        to_char(orde.order_id) 编号,\n" +
"                        rad.jy_region_name 区域,\n" +
"                        busi.business_name 受理类型,\n" +
"                        orde.create_date 时间,\n" +
"                        sta.staff_name 操作员,\n" +
"                        case\n" +
"                          when org.organize_name like '%营业厅%' then\n" +
"                           (select org1.organize_name\n" +
"                              from params2.sec_organize org1\n" +
"                             where org1.organize_id = org.parent_organize_id)\n" +
"                          else\n" +
"                           org.organize_name\n" +
"                        end 所属广电站\n" +
"          from cp2.cm_customer cust,\n" +
"               cp2.cb_party part,\n" +
"               cp2.cb_party_role pole,\n" +
"               (select *\n" +
"                  from jour2.om_order\n" +
"                union all\n" +
"                select *\n" +
"                  from jour2.om_order_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + ") orde,\n" +
"               pzg1.bs_business busi,\n" +
"               params1.sec_operator sop,\n" +
"               params1.sec_staff sta,\n" +
"               params1.sec_organize org,\n" +
"               wxjy.jy_region_address_rel rad\n" +
"         where cust.party_id = part.party_id\n" +
"           and cust.partition_id = pole.partition_id\n" +
"           and cust.party_id = pole.party_id\n" +
"           and orde.party_role_id = pole.party_role_id\n" +
"           and orde.busi_code = to_char(busi.business_type_id)\n" +
"           and orde.op_id = sop.operator_id\n" +
"           and sop.staff_id = sta.staff_id\n" +
"           and orde.org_id = org.organize_id\n" +
"           and cust.cust_id = rad.cust_id\n" +
"           and org.organize_name <> '系统自动任务'\n" +
"           and orde.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"           and orde.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"           and cust.own_corp_org_id = 3328\n" +
"        union all\n" +
"        select distinct cust.cust_code 客户证号,\n" +
"                        part.party_name 姓名,\n" +
"                        to_char(pay.business_id) 编号,\n" +
"                        rad.jy_region_name 区域,\n" +
"                        busi.business_name 受理类型,\n" +
"                        pay.create_date 时间,\n" +
"                        sta.staff_name 操作员,\n" +
"                        case\n" +
"                          when org.organize_name like '%营业厅%' then\n" +
"                           (select org1.organize_name\n" +
"                              from params2.sec_organize org1\n" +
"                             where org1.organize_id = org.parent_organize_id)\n" +
"                          else\n" +
"                           org.organize_name\n" +
"                        end 所属广电站\n" +
"          from cp2.cm_customer            cust,\n" +
"               cp2.cb_party               part,\n" +
"               files2.cm_account          acct,\n" +
"               ac2.am_busi                pay,\n" +
"               pzg1.bs_business           busi,\n" +
"               params1.sec_operator       sop,\n" +
"               params1.sec_staff          sta,\n" +
"               params1.sec_organize       org,\n" +
"               wxjy.jy_region_address_rel rad\n" +
"         where cust.party_id = part.party_id\n" +
"           and cust.cust_id = acct.cust_id\n" +
"           and acct.acct_id = pay.acct_id\n" +
"           and pay.trade_op_id = sop.operator_id\n" +
"           and sop.staff_id = sta.staff_id\n" +
"           and sta.organize_id = org.organize_id\n" +
"           and cust.cust_id = rad.cust_id\n" +
"           and pay.business_type_id = busi.business_type_id\n" +
"           and pay.amount <> 0\n" +
"           and pay.cancel_flag = 'U' ---正常缴费未返销\n" +
"           and sta.staff_name <> 'UPG'\n" +
"           and pay.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"           and pay.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"           and cust.own_corp_org_id = 3328) tmp,\n" +
"       wxjy.jy_organize_region_rel rel\n" +
" where tmp.所属广电站 = rel.organize_name\n" +
"   and tmp.区域 <> rel.region_name\n" +
" order by tmp.操作员, tmp.时间, tmp.客户证号, tmp.编号";

                int[] columntxt = { 1, 3 }; //哪些列是文本格式
                int[] columndate = { 6 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "营业员异地受理明细数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报续费返充即将到期的客户明细
        public static bool ybaobiao32()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "续费返充即将到期的客户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct prod.offer_name 产品,\n" +
"                aapl.origin_balance / 100 返还总额_元,\n" +
"                aapl.months 返还总次数,\n" +
"                aapl.start_month 开始返还账期,\n" +
"                aapl.remain_months 剩余次数,\n" +
"                aapl.month_return_amount / 100 每次返还金额_元,\n" +
"                aapl.remain_balance / 100 剩余金额_元,\n" +
"                cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_address          addr,\n" +
"       files2.um_subscriber       subs,\n" +
"       ac2.am_apportion           appo,\n" +
"       ac2.am_apportion_plan      aapl,\n" +
"       upc1.pm_offer              prod,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = appo.subscriber_id\n" +
"   and appo.price_ins_id = aapl.price_ins_id\n" +
"   and aapl.offer_id = prod.offer_id\n" +
"   and addr.expire_date > sysdate\n" +
"   and aapl.month_return_amount > 0\n" +
"   and aapl.state = 'R'\n" +
"   and aapl.remain_months = 1\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by prod.offer_name, cust.cust_code";

                int[] columntxt = { 8, 11 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "续费返充即将到期的客户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报上2个月欠费停机主动暂停用户明细
        public static bool ybaobiao33()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "上2个月欠费停机主动暂停用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select *\n" +
"  from (select distinct cust.cust_code 客户证号,\n" +
"                        part.party_name 姓名,\n" +
"                        decode(cust.cust_type,\n" +
"                               1,\n" +
"                               '公众客户',\n" +
"                               2,\n" +
"                               '商业客户',\n" +
"                               3,\n" +
"                               '团体代付客户',\n" +
"                               4,\n" +
"                               '合同商业客户') 客户类型,\n" +
"                        term.serial_no 机顶盒,\n" +
"                        subs.sub_bill_id 智能卡,\n" +
"                        case\n" +
"                          when nvl(subs.main_subscriber_ins_id, 0) = 0 then\n" +
"                           '主机'\n" +
"                          else\n" +
"                           ''\n" +
"                        end 主副机,\n" +
"                        case\n" +
"                          when subs.state = 'M' then\n" +
"                           '暂停'\n" +
"                          when fsta.os_status = '1' then\n" +
"                           '欠费停机'\n" +
"                          when fsta.os_status in ('3', '4') then\n" +
"                           '暂停'\n" +
"                          else\n" +
"                           ''\n" +
"                        end 停开机状态,\n" +
"                        case\n" +
"                          when subs.state = 'M' then\n" +
"                           subs.done_date\n" +
"                          when fsta.os_status = '1' then\n" +
"                           fsta.done_date\n" +
"                          when fsta.os_status in (3, 4) then\n" +
"                           subs.done_date\n" +
"                          else\n" +
"                           null\n" +
"                        end 受理时间,\n" +
"                        rsku.res_sku_name 资源类型,\n" +
"                        rel.teleph_nunber 联系方式,\n" +
"                        rad.jy_region_name 区域,\n" +
"                        addr.std_addr_name 地址,\n" +
"                        wg.region_name 网格区域,\n" +
"                        wg.grid_name 所属网格\n" +
"          from cp2.cm_customer cust,\n" +
"               cp2.cb_party               part,\n" +
"               files2.um_subscriber       subs,\n" +
"               files2.um_offer_06         ofer,\n" +
"               files2.um_offer_sta_02     fsta,\n" +
"               files2.um_res              ures,\n" +
"               res1.res_terminal          term,\n" +
"               res1.res_sku               rsku,\n" +
"               files2.um_address          addr,\n" +
"               wxjy.jy_region_address_rel rad,\n" +
"               wxjy.jy_contact_rel        rel,\n" +
"               wxjy.jy_customer_wg_rel    wg\n" +
"         where cust.party_id = part.party_id\n" +
"           and cust.cust_id = subs.cust_id\n" +
"           and subs.subscriber_ins_id = ures.subscriber_ins_id\n" +
"           and ures.res_equ_no = term.serial_no\n" +
"           and term.res_sku_id = rsku.res_sku_id\n" +
"           and cust.cust_id = addr.cust_id(+)\n" +
"           and cust.cust_id = rad.cust_id(+)\n" +
"           and cust.party_id = rel.party_id\n" +
"           and cust.partition_id = rel.partition_id\n" +
"           and cust.cust_id = wg.cust_id(+)\n" +
"           and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"           and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"           and ofer.prod_service_id = 1002 --- 基本节目\n" +
"           and subs.main_spec_id = 80020001 ---数字电视\n" +
"           and addr.expire_date > sysdate\n" +
"           and ures.expire_date > sysdate\n" +
"           and fsta.expire_date > sysdate\n" +
"           and ofer.expire_date > sysdate\n" +
"           and rsku.res_type_id = 2\n" +
"           and cust.own_corp_org_id = 3328\n" +
"           and not exists\n" +
"         (select 1\n" +
"                  from files2.um_offer_06 ofer1, files2.um_offer_sta_02 fsta1\n" +
"                 where ofer1.offer_ins_id = fsta1.offer_ins_id\n" +
"                   and ofer1.prod_service_id = 1002\n" +
"                   and fsta1.offer_status = '1'\n" +
"                   and fsta1.os_status is null\n" +
"                   and fsta1.expire_date > sysdate\n" +
"                   and ofer1.expire_date > sysdate\n" +
"                   and ofer1.subscriber_ins_id = subs.subscriber_ins_id)) tmp\n" +
" where tmp.受理时间 >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-3).ToString("yyyy-MM-dd") + "'\n" +
"   and tmp.受理时间 < date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
" order by tmp.区域, tmp.客户证号, tmp.机顶盒";

                int[] columntxt = { 1, 4, 5, 10 }; //哪些列是文本格式
                int[] columndate = { 8 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "上2个月欠费停机主动暂停用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报高清互动休眠用户(未开通点播功能)明细
        public static bool ybaobiao34()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "高清互动休眠用户(未开通点播功能)明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                term.serial_no 机顶盒,\n" +
"                subs.sub_bill_id 智能卡,\n" +
"                case\n" +
"                  when nvl(subs.main_subscriber_ins_id, 0) = 0 then\n" +
"                   '主机'\n" +
"                  else\n" +
"                   ''\n" +
"                end 主副机,\n" +
"                rsku.res_sku_name 资源类型,\n" +
"                ures.create_date 资源订购时间,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_res              ures,\n" +
"       res1.res_terminal          term,\n" +
"       res1.res_sku               rsku,\n" +
"       files2.um_address          addr,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_customer_wg_rel    wg\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ures.subscriber_ins_id\n" +
"   and ures.res_equ_no = term.serial_no\n" +
"   and term.res_sku_id = rsku.res_sku_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and addr.expire_date > sysdate\n" +
"   and ures.expire_date > sysdate\n" +
"   and rsku.res_type_id = 2\n" +
"   and cust.cust_type = 1 ---只统计公众客户\n" +
"   and cust.cust_prop <> '6' ---不统计免费客户\n" +
"   and part.party_name not like '%ceshi%'\n" +
"   and part.party_name not like '%测试%'\n" +
"   and part.party_name not like '%test%'\n" +
"   and rsku.res_sku_name in ('银河高清基本型HDC6910(江阴)',\n" +
"                             '银河高清交互型HDC691033(江阴)',\n" +
"                             '银河智能高清交互型HDC6910798(江阴)',\n" +
"                             '银河4K交互型HDC691090',\n" +
"                             '4K超高清型II型融合型（EOC）',\n" +
"                             '银河4K超高清II型融合型HDC6910B1(江阴)',\n" +
"                             '4K超高清简易型（基本型）')\n" +
"   and cust.own_corp_org_id = 3328\n" +
"   and exists\n" +
" (select 1\n" +
"          from files2.um_offer_06 ofer, files2.um_offer_sta_02 fsta\n" +
"         where ofer.offer_ins_id = fsta.offer_ins_id\n" +
"           and ofer.expire_date > sysdate\n" +
"           and fsta.expire_date > sysdate\n" +
"           and ofer.prod_service_id = 1002\n" +
"           and fsta.offer_status = '1'\n" +
"           and fsta.os_status is null\n" +
"           and ofer.subscriber_ins_id = subs.subscriber_ins_id) ---有正常的基本节目\n" +
"   and not exists\n" +
" (select 1\n" +
"          from files2.um_offer_06     ofer1,\n" +
"               files2.um_offer_sta_02 fsta1,\n" +
"               wxjy.jy_dbitv_product  jitv1\n" +
"         where ofer1.offer_ins_id = fsta1.offer_ins_id\n" +
"           and fsta1.offer_id = jitv1.offer_id\n" +
"           and ofer1.expire_date > sysdate\n" +
"           and fsta1.expire_date > sysdate\n" +
"           and fsta1.offer_status = '1'\n" +
"           and fsta1.os_status is null\n" +
"           and ofer1.subscriber_ins_id = subs.subscriber_ins_id) ---没有正常的互动产品\n" +
"   and not exists\n" +
" (select 1\n" +
"          from files2.um_subscriber   subs2,\n" +
"               files2.um_offer_06     ofer2,\n" +
"               files2.um_offer_sta_02 fsta2,\n" +
"               wxjy.jy_dbitv_product  jitv2\n" +
"         where subs2.subscriber_ins_id = ofer2.subscriber_ins_id\n" +
"           and ofer2.offer_ins_id = fsta2.offer_ins_id\n" +
"           and fsta2.offer_id = jitv2.offer_id\n" +
"           and subs2.main_spec_id = 80020199\n" +
"           and ofer2.expire_date > sysdate\n" +
"           and fsta2.expire_date > sysdate\n" +
"           and subs2.cust_id = cust.cust_id) ---客户下不存在有订购正常互动产品的虚用户（即客户级订购）\n" +
" order by cust.cust_code, term.serial_no";

                int[] columntxt = { 1, 4, 5 }; //哪些列是文本格式
                int[] columndate = { 8 };         //哪些列是日期格式
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                ExcelHelper.DataTableToExcel("\\月报\\" + "高清互动休眠用户(未开通点播功能)明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报各业务分账本销账金额汇总数据（最新）
        public static bool ybaobiao36()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "各业务分账本销账金额汇总数据（最新）--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select * from (\n" +
"select rad.jy_region_name 区域,\n" +
"       decode(cust.cust_type,\n" +
"              1,\n" +
"              '公众客户',\n" +
"              2,\n" +
"              '商业客户',\n" +
"              3,\n" +
"              '团体代付客户',\n" +
"              4,\n" +
"              '合同商业客户') 客户类型,\n" +
"       decode(a.service_id, 1002, '数字基本', 1003, '互动基本', 1004, '宽带', 1005, '付费节目', 1006, '互动点播', 1008, '增值') 分业务,\n" +
"       sum(case when a.asset_item_id = 100 then a.writeoff_fee else 0 end) / 100 通用现金账本,\n" +
"       sum(case when a.asset_item_id = 200 then a.writeoff_fee else 0 end) / 100 数字基本现金账本,\n" +
"       sum(case when a.asset_item_id = 300 then a.writeoff_fee else 0 end) / 100 互动基本现金账本,\n" +
"       sum(case when a.asset_item_id = 400 then a.writeoff_fee else 0 end) / 100 专用账本,\n" +
"       sum(case when a.asset_item_id = 500 then a.writeoff_fee else 0 end) / 100 划转账本,\n" +
"       sum(case when a.asset_item_id = 600 then a.writeoff_fee else 0 end) / 100 虚拟通用账本,\n" +
"       sum(case when a.asset_item_id = 700 then a.writeoff_fee else 0 end) / 100 宽带现金账本,\n" +
"       sum(case when a.asset_item_id not in (100,200,300,400,500,600,700) then a.writeoff_fee else 0 end) / 100 其他账本\n" +
"  from rep.fin2_writeoff_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "   a,\n" +
"       rep.fin2_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "      cust,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where a.acct_id = cust.acct_id\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and a.corp_org_id = 3328\n" +
"   and a.bill_item_id not in (22012, 22013, 22014, 22017) ---2023年3月3号修改，剔除挂账退费的账单\n" +
"   and not ((cust.cust_type = 1 and cust.cust_prop = 6) or (cust.cust_type = 2 and cust.cust_prop in (8, 10)) or (cust.cust_type = 4 and cust.cust_prop = 10))\n" +
" group by rad.jy_region_name, cust.cust_type, a.service_id) tmp\n" +
" order by tmp.区域, tmp.客户类型, tmp.分业务";
                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 0 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\月报\\" + "各业务分账本销账金额汇总数据（最新）--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报各业务分账本销账金额汇总数据（最新，区域为空明细）
        public static bool ybaobiao37()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "各业务分账本销账金额汇总数据（最新，区域为空明细）--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select aa.*\n" +
"  from (select cust.cust_code 客户证号,\n" +
"               cust.acct_name 姓名,\n" +
"               decode(cust.cust_type,\n" +
"                      1,\n" +
"                      '公众客户',\n" +
"                      2,\n" +
"                      '商业客户',\n" +
"                      3,\n" +
"                      '团体代付客户',\n" +
"                      4,\n" +
"                      '合同商业客户') 客户类型,\n" +
"               rad.jy_region_name 区域,\n" +
"               cust.stand_name 标准地址,\n" +
"       decode(a.service_id, 1002, '数字基本', 1003, '互动基本', 1004, '宽带', 1005, '付费节目', 1006, '互动点播', 1008, '增值') 分业务,\n" +
"       sum(case when a.asset_item_id = 100 then a.writeoff_fee else 0 end) / 100 通用现金账本,\n" +
"       sum(case when a.asset_item_id = 200 then a.writeoff_fee else 0 end) / 100 数字基本现金账本,\n" +
"       sum(case when a.asset_item_id = 300 then a.writeoff_fee else 0 end) / 100 互动基本现金账本,\n" +
"       sum(case when a.asset_item_id = 400 then a.writeoff_fee else 0 end) / 100 专用账本,\n" +
"       sum(case when a.asset_item_id = 500 then a.writeoff_fee else 0 end) / 100 划转账本,\n" +
"       sum(case when a.asset_item_id = 600 then a.writeoff_fee else 0 end) / 100 虚拟通用账本,\n" +
"       sum(case when a.asset_item_id = 700 then a.writeoff_fee else 0 end) / 100 宽带现金账本,\n" +
"       sum(case when a.asset_item_id not in (100,200,300,400,500,600,700) then a.writeoff_fee else 0 end) / 100 其他账本\n" +
"          from rep.fin2_writeoff_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "   a,\n" +
"               rep.fin2_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "      cust,\n" +
"               wxjy.jy_region_address_rel rad\n" +
"         where a.acct_id = cust.acct_id\n" +
"           and cust.cust_id = rad.cust_id(+)\n" +
"           and a.corp_org_id = 3328\n" +
"           and a.bill_item_id not in (22012, 22013, 22014, 22017) ---2023年3月3号修改，剔除挂账退费的账单\n" +
"           and not ((cust.cust_type = 1 and cust.cust_prop = 6) or\n" +
"                     (cust.cust_type = 2 and cust.cust_prop = 8))\n" +
" group by rad.jy_region_name, cust.cust_type, a.service_id, cust.cust_code, cust.acct_name, cust.stand_name) aa\n" +
" where aa.区域 is null\n" +
" order by aa.标准地址";

                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 0 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\月报\\" + "各业务分账本销账金额汇总数据（最新，区域为空明细）--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报感恩套餐即将到期的客户明细数据
        public static bool ybaobiao38()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "感恩套餐即将到期的客户明细数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                prod.offer_name 套餐,\n" +
"                ofer.create_date 订购时间,\n" +
"                ofer.valid_date 生效时间,\n" +
"                ofer.expire_date 失效时间,\n" +
"                case\n" +
"                  when exists (select 1\n" +
"                          from ac2.am_entrust_relation rela\n" +
"                         where rela.exp_date > sysdate\n" +
"                           and rela.state = 0\n" +
"                           and rela.acct_id = acct.acct_id) then\n" +
"                   '银行托收'\n" +
"                  else\n" +
"                   ''\n" +
"                end 缴费方式,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.cm_account          acct,\n" +
"       files2.um_address          addr,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_offer_06         ofer,\n" +
"       upc1.pm_offer              prod,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = acct.cust_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.offer_id = prod.offer_id\n" +
"   and addr.expire_date > sysdate\n" +
"   and ofer.expire_date > sysdate\n" +
"   and cust.cust_type = 1\n" +
"   and cust.cust_prop <> '6'\n" +
"   and ofer.offer_name in ('高清幸福感恩套餐',\n" +
"                           '高清高速感恩套餐',\n" +
"                           '高清极速感恩套餐',\n" +
"                           '幸福感恩置换',\n" +
"                           '518感恩套餐',\n" +
"                           '高速感恩置换',\n" +
"                           '988感恩套餐',\n" +
"                           '极速感恩置换',\n" +
"                           '788感恩套餐',\n" +
"                           '899优生活套餐')\n" +
"   and ((ofer.valid_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+1).ToString("2025-MM-01") + "' and ofer.valid_date < date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+2).ToString("2025-MM-01") + "')\n" +
"     or (ofer.valid_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+1).ToString("2024-MM-01") + "' and ofer.valid_date < date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+2).ToString("2024-MM-01") + "')\n" +
"     or (ofer.valid_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+1).ToString("2023-MM-01") + "' and ofer.valid_date < date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+2).ToString("2023-MM-01") + "')\n" +
"     or (ofer.valid_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+1).ToString("2022-MM-01") + "' and ofer.valid_date < date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+2).ToString("2022-MM-01") + "')\n" +
"     or (ofer.valid_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+1).ToString("2021-MM-01") + "' and ofer.valid_date < date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+2).ToString("2021-MM-01") + "')\n" +
"     or (ofer.valid_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+1).ToString("2020-MM-01") + "' and ofer.valid_date < date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+2).ToString("2020-MM-01") + "')\n" +
"     or (ofer.valid_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+1).ToString("2019-MM-01") + "' and ofer.valid_date < date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+2).ToString("2019-MM-01") + "')\n" +
"     or (ofer.valid_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+1).ToString("2018-MM-01") + "' and ofer.valid_date < date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+2).ToString("2018-MM-01") + "')\n" +
"     or (ofer.valid_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+1).ToString("2017-MM-01") + "' and ofer.valid_date < date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+2).ToString("2017-MM-01") + "')\n" +
"     or (ofer.valid_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+1).ToString("2016-MM-01") + "' and ofer.valid_date < date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+2).ToString("2016-MM-01") + "')\n" +
"     or (ofer.valid_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+1).ToString("2015-MM-01") + "' and ofer.valid_date < date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+2).ToString("2015-MM-01") + "')\n" +
"     or (ofer.valid_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+1).ToString("2014-MM-01") + "' and ofer.valid_date < date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")).AddMonths(+2).ToString("2014-MM-01") + "'))\n" +
"   and cust.own_corp_org_id = 3328\n" +
"   and exists (select 1\n" +
"          from files2.um_offer_sta_02 fsta\n" +
"         where fsta.expire_date > sysdate\n" +
"           and fsta.offer_status = '1'\n" +
"           and fsta.os_status is null\n" +
"           and fsta.offer_ins_id = ofer.offer_ins_id)\n" +
"   and exists\n" +
" (select 1\n" +
"          from files2.um_offer_06 ofer1, files2.um_offer_sta_02 fsta1\n" +
"         where ofer1.offer_ins_id = fsta1.offer_ins_id\n" +
"           and ofer1.expire_date > sysdate\n" +
"           and fsta1.expire_date > sysdate\n" +
"           and ofer1.prod_service_id in (1002, 1004)\n" +
"           and fsta1.offer_status = '1'\n" +
"           and fsta1.os_status is null\n" +
"           and ofer1.subscriber_ins_id = subs.subscriber_ins_id)\n" +
" order by cust.cust_code, ofer.create_date";


                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 1, 9 }; //哪些列是文本格式
                int[] columndate = { 5, 6, 7 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\月报\\" + "感恩套餐即将到期的客户明细数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报集团出销账明细
        public static bool ybaobiao39()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "集团出销账明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =

"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "' 月份,\n" +
"                a1.amount1 出账金额,\n" +
"                a2.amont2 销账金额\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.cm_account          acct,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       files2.um_address          addr,\n" +
"       (select cz.acct_id, sum(cz.fee) / 100 amount1\n" +
"          from ac2.am_bill_ar_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " cz\n" +
"         group by cz.acct_id) a1,\n" +
"       (select xz.acct_id, sum(xz.writeoff_fee) / 100 amont2\n" +
"          from (select pxz.acct_id, pxz.writeoff_fee\n" +
"                  from ac2.am_writeoff_d pxz\n" +
"                 where pxz.bill_month = " + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " \n" +
"                   and pxz.bill_flag = 'U'\n" +
"                union all\n" +
"                select cxz.acct_id, cxz.writeoff_fee\n" +
"                  from ac2.am_writeoff cxz\n" +
"                 where cxz.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"                   and cxz.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"                   and cxz.bill_flag = 'U') xz\n" +
"         group by xz.acct_id) a2\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = acct.cust_id\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and acct.acct_id = a1.acct_id(+)\n" +
"   and acct.acct_id = a2.acct_id(+)\n" +
"   and (a1.amount1 <> 0 or a2.amont2 <> 0)\n" +
"   and addr.expire_date > sysdate\n" +
"   and cust.cust_type = 4\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name, cust.cust_code";

                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 1 }; //哪些列是文本格式
                int[] columndate = { 0 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\月报\\" + "集团出销账明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报5月欠费客户的当前状态
        public static bool ybaobiao40()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "5月欠费客户的当前状态--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                 jcq.done_date 历史欠费时间,\n" +
"                 jcq.fee 历史账单欠费金额,\n" +
"                 a.status 当前状态,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址\n" +
"  from cp2.cm_customer            cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.cm_account          acct,\n" +
"       files2.um_address          addr,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_offer_06         ofer,\n" +
"       wxjy.jy_cust_qianfei_0510     jcq,\n" +
"        wxjy.jy_customer_status a\n" +
" where cust.party_id = part.party_id\n" +
" and jcq.cust_code = a.cust_code\n" +
"   and  jcq.cust_code = cust.cust_code\n" +
"   and cust.cust_id = acct.cust_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and nvl(subs.main_subscriber_ins_id, 0) = 0\n" +
"   and ofer.prod_service_id = 1002\n" +
"   and addr.expire_date > sysdate\n" +
"   and ofer.expire_date > sysdate\n" +
"   and cust.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name, addr.std_addr_name, cust.cust_code";

                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 1,7 }; //哪些列是文本格式
                int[] columndate = { 4 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\月报\\" + "5月欠费客户的当前状态--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion

        #region 月报欠费停机以及暂停用户明细
        public static bool ybaobiao41()
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            if (File.Exists(pathInfo.Parent.FullName + "\\月报\\" + "欠费停机以及暂停用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx")) { return true; }
            else
            {
                string sqlString =
"select distinct cust.cust_code 客户证号,\n" +
"                part.party_name 姓名,\n" +
"                decode(cust.cust_type,\n" +
"                       1,\n" +
"                       '公众客户',\n" +
"                       2,\n" +
"                       '商业客户',\n" +
"                       3,\n" +
"                       '团体代付客户',\n" +
"                       4,\n" +
"                       '合同商业客户') 客户类型,\n" +
"                term.serial_no 机顶盒,\n" +
"                subs.sub_bill_id 智能卡,\n" +
"case\n" +
"            when exists (select 1\n" +
"                    from files2.um_offer_06     ofer1,\n" +
"                         files2.um_offer_sta_02 fsta1,\n" +
"                         wxjy.jy_dbitv_product  hd\n" +
"                   where ofer1.offer_ins_id = fsta1.offer_ins_id\n" +
"                     and fsta1.offer_id = hd.offer_id\n" +
"                     and ofer1.subscriber_ins_id =\n" +
"                         subs.subscriber_ins_id) then\n" +
"             '有'\n" +
"            else\n" +
"             ''\n" +
"          end 是否有互动, \n" +
"                case\n" +
"                  when nvl(subs.main_subscriber_ins_id, 0) = 0 then\n" +
"                   '主机'\n" +
"                  else\n" +
"                   ''\n" +
"                end 主副机,\n" +
"                case\n" +
"                  when subs.state = 'M' then\n" +
"                   '暂停'\n" +
"                  when fsta.os_status = '1' then\n" +
"                   '欠费停机'\n" +
"                  when fsta.os_status in ('3', '4') then\n" +
"                   '暂停'\n" +
"                  else\n" +
"                   ''\n" +
"                end 停开机状态,\n" +
"                case\n" +
"                  when subs.state = 'M' then\n" +
"                   subs.done_date\n" +
"                  when fsta.os_status = '1' then\n" +
"                   fsta.done_date\n" +
"                  when fsta.os_status in (3, 4) then\n" +
"                   subs.done_date\n" +
"                  else\n" +
"                   null\n" +
"                end 受理时间,\n" +
"                rsku.res_sku_name 资源类型,\n" +
"                rel.teleph_nunber 联系方式,\n" +
"                rad.jy_region_name 区域,\n" +
"                addr.std_addr_name 地址,\n" +
"                wg.region_name 网格区域,\n" +
"                wg.grid_name 所属网格\n" +
"  from cp2.cm_customer cust,\n" +
"       cp2.cb_party               part,\n" +
"       files2.um_subscriber       subs,\n" +
"       files2.um_offer_06         ofer,\n" +
"       files2.um_offer_sta_02     fsta,\n" +
"       files2.um_res              ures,\n" +
"       res1.res_terminal          term,\n" +
"       res1.res_sku               rsku,\n" +
"       files2.um_address          addr,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_contact_rel        rel,\n" +
"       wxjy.jy_customer_wg_rel    wg\n" +
" where cust.party_id = part.party_id\n" +
"   and cust.cust_id = subs.cust_id\n" +
"   and subs.subscriber_ins_id = ures.subscriber_ins_id\n" +
"   and ures.res_equ_no = term.serial_no\n" +
"   and term.res_sku_id = rsku.res_sku_id\n" +
"   and cust.cust_id = addr.cust_id(+)\n" +
"   and cust.cust_id = rad.cust_id(+)\n" +
"   and cust.party_id = rel.party_id\n" +
"   and cust.partition_id = rel.partition_id\n" +
"   and cust.cust_id = wg.cust_id(+)\n" +
"   and subs.subscriber_ins_id = ofer.subscriber_ins_id\n" +
"   and ofer.offer_ins_id = fsta.offer_ins_id\n" +
"   and ofer.prod_service_id = 1002 --- 基本节目\n" +
"   and subs.main_spec_id = 80020001 ---数字电视\n" +
"   and addr.expire_date > sysdate\n" +
"   and ures.expire_date > sysdate\n" +
"   and fsta.expire_date > sysdate\n" +
"   and ofer.expire_date > sysdate\n" +
"   and rsku.res_type_id = 2\n" +
"   and cust.own_corp_org_id = 3328\n" +
"   and not exists\n" +
" (select 1\n" +
"          from files2.um_offer_06 ofer1, files2.um_offer_sta_02 fsta1\n" +
"         where ofer1.offer_ins_id = fsta1.offer_ins_id\n" +
"           and ofer1.prod_service_id = 1002\n" +
"           and fsta1.offer_status = '1'\n" +
"           and fsta1.os_status is null\n" +
"           and fsta1.expire_date > sysdate\n" +
"           and ofer1.expire_date > sysdate\n" +
"           and ofer1.subscriber_ins_id = subs.subscriber_ins_id)\n" +
" order by rad.jy_region_name,\n" +
"          addr.std_addr_name,\n" +
"          cust.cust_code,\n" +
"          term.serial_no";

                DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
                int[] columntxt = { 1, 4,5,10 }; //哪些列是文本格式
                int[] columndate = { 9 };         //哪些列是日期格式
                ExcelHelper.DataTableToExcel("\\月报\\" + "欠费停机以及暂停用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
                return true;
            }
        }
        #endregion
    }
}
