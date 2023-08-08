using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace AutoUpDataBoss
{
     class DataProcessing
    {
          #region 创建复通所需临时表
        public static bool cbaobiao1()
        {
            string sqlString =
"create table wxjy.jy_tjkh_"+ DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).AddDays(-3).ToString("yyyyMMdd").ToString() + " as\n" +
"select scc.cust_id, scc.cust_code, scc.cust_name\n" +
"  from so1.cm_customer scc\n" +
" where scc.own_corp_org_id = 3328\n" +
"   and exists (select 1\n" +
"          from repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-3).AddDays(-1).ToString("yyyyMMdd").ToString() + " sip\n" +
"         where sip.prod_spec_id = 800200000001\n" +
"           and sip.sub_bill_id is not null\n" +
"           and scc.cust_id = sip.cust_id) --存在数字电视用户\n" +
"   and not exists (select 1\n" +
"          from repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-3).AddDays(-1).ToString("yyyyMMdd").ToString() + "   sip1,\n" +
"               repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-3).AddDays(-1).ToString("yyyyMMdd").ToString() + " sis1\n" +
"         where sip1.prod_inst_id = sis1.prod_inst_id\n" +
"           and sis1.prod_service_id = 1002\n" +
"           and sis1.state = '1'\n" +
"           and sis1.os_status is null\n" +
"           and sip1.cust_id = scc.cust_id) --不存在正常的数字电视用户";
            return OracleHelper.ExecuteNonQuery(sqlString);
        }
        #endregion
        #region 创建复通所需临时表后计数
        public static string cbaobiao2()
        {
            string sqlString =
            "select count(*) cou from jy_tjkh_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-3).AddDays(-1).ToString("yyyyMMdd").ToString();
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            string count1 = dt.Rows[0]["cou"].ToString();
            return count1;
        }
        #endregion








        #region 日报订购产品
        public static bool baobiao1()
        {
            string oneday = "trunc(sysdate)-1";
            if(DateTime.Now.DayOfWeek.ToString()== "Monday")
            { oneday = "trunc(sysdate)-3"; }
            string sqlString =
"select scc.cust_code 客户证号,\n" +
"       scc.cust_name 客户姓名,\n" +
"       ipr.res_equ_no 机顶盒号码,\n" +
"       sip.sub_bill_id 智能卡,\n" +
"       upd.name 产品名称,\n" +
"       sis.create_date 订购时间,\n" +
"       sor.organize_name 营业厅,\n" +
"       sta.staff_name 操作员,\n" +
"       con.cont_phone1 || ',' || con.cont_phone2 联系方式,\n" +
"       con.cont_mobile1 || ',' || con.cont_mobile2 移动电话,\n" +
"       rad.jy_region_name 区域,\n" +
"       iad.std_addr_name || iad.door_name 地址\n" +
"  from so1.cm_customer            scc,\n" +
"       so1.ins_prod               sip,\n" +
"       so1.ins_srvpkg             sis,\n" +
"       product.up_product_item    upd,\n" +
"       so1.ins_prod_res           ipr,\n" +
"       sec.sec_operator           sop,\n" +
"       sec.sec_staff              sta,\n" +
"       so1.ins_address            iad,\n" +
"       so1.cm_cust_contact_info   con,\n" +
"       sec.sec_organize           sor,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where scc.cust_id = sip.cust_id\n" +
"   and sip.prod_inst_id = sis.prod_inst_id\n" +
"   and sis.srvpkg_id = upd.product_item_id\n" +
"   and ipr.prod_inst_id = sip.prod_inst_id\n" +
"   and sis.op_id = sop.operator_id\n" +
"   and sta.staff_id = sop.staff_id\n" +
"   and scc.cust_id = iad.cust_id\n" +
"   and scc.cust_id = con.cust_id\n" +
"   and sta.organize_id = sor.organize_id\n" +
"   and scc.cust_id = rad.cust_id\n" +
"   and （upd.name like '%炫力动漫%' or upd.name like '%学霸宝盒OTT%'）\n" +
"   and sis.state = 1\n" +
"   and sis.os_status is null\n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and sis.create_date >= "+ oneday + "\n" +
"   and sis.create_date < trunc(sysdate)\n" +
"   and exists (select 1\n" +
"          from res.res_code_definition rcd\n" +
"         where rcd.res_type = 2\n" +
"           and ipr.res_code = rcd.res_code)\n" +
" order by sis.create_date, scc.cust_code";
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            int[] columntxt = {1,3,4,9,10}; //哪些列是文本格式
            int[] columndate = {6};         //哪些列是日期格式
            ExcelHelper.DataTableToExcel("\\日报\\" + "日报订购产品--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 日报退订产品
        public static bool baobiao2()
        {
            string oneday = "trunc(sysdate)-1";
            if (DateTime.Now.DayOfWeek.ToString() == "Monday")
            { oneday = "trunc(sysdate)-3"; }
            string sqlString =
"select distinct scc.cust_code      客户证号,\n" +
"                scc.cust_name      客户姓名,\n" +
"                prod.bill_id       机顶盒号码,\n" +
"                upd.name           产品名称,\n" +
"                inof.create_date   订购时间,\n" +
"                spkg.create_date   退订时间,\n" +
"                sor.organize_name  营业厅,\n" +
"                sta.staff_name     操作员,\n" +
"                rad.jy_region_name 区域,\n" +
"                iad.std_addr_name  地址\n" +
"  from so1.cm_customer scc,\n" +
"       (select *\n" +
"          from so1.ord_cust\n" +
"        union all\n" +
"        select *\n" +
"          from so1.ord_cust_f_2020) cust,\n" +
"       (select *\n" +
"          from so1.ord_prod\n" +
"        union all\n" +
"        select *\n" +
"          from so1.ord_prod_f_2020) prod,\n" +
"       (select *\n" +
"          from so1.ord_srvpkg\n" +
"        union all\n" +
"        select *\n" +
"          from so1.ord_srvpkg_f_2020) spkg,\n" +
"       (select *\n" +
"          from so1.ord_offer\n" +
"        union all\n" +
"        select *\n" +
"          from so1.ord_offer_f_2020) offe,\n" +
"       (select * from so1.h_ins_offer_2020) inof,\n" +
"       product.up_product_item upd,\n" +
"       so1.ins_address iad,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       sec.sec_operator sop,\n" +
"       sec.sec_staff sta,\n" +
"       sec.sec_organize sor\n" +
" where scc.cust_id = cust.cust_id\n" +
"   and scc.cust_id = iad.cust_id\n" +
"   and scc.cust_id = rad.cust_id\n" +
"   and cust.cust_order_id = prod.cust_order_id\n" +
"   and prod.prod_order_id = spkg.prod_order_id\n" +
"   and offe.offer_order_id = spkg.offer_order_id\n" +
"   and offe.offer_inst_id = inof.offer_inst_id\n" +
"   and spkg.srvpkg_id = upd.product_item_id\n" +
"   and cust.op_id = sop.operator_id(+)\n" +
"   and sop.staff_id = sta.staff_id(+)\n" +
"   and sta.organize_id = sor.organize_id(+)\n" +
"   and （upd.name like '%炫力动漫%' or upd.name like '%学霸宝盒OTT%'）\n" +
"   and cust.business_id = 800001000026\n" +
"   and spkg.state = 3\n" +
"   and spkg.create_date >= " + oneday + "\n" +
"   and spkg.create_date < trunc(sysdate)\n" +
"   and scc.own_corp_org_id = 3328\n" +
" order by spkg.create_date";
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            int[] columntxt = { 1, 3 }; //哪些列是文本格式
            int[] columndate = { 5,6 };         //哪些列是日期格式
            ExcelHelper.DataTableToExcel("\\日报\\" + "日报退订产品--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 日报新增电视客户明细
        public static bool baobiao4()
        {
            string sqlString =

"select tp.区域,\n" +
"       trunc(sysdate) - 1 日期,\n" +
"        tp1.客户类型,\n" +
"       nvl(tp1.新增数字电视客户数, 0) 新增数字电视客户数\n" +
"  from (select distinct rad.jy_region_name 区域\n" +
"          from wxjy.jy_region_address_rel rad) tp,\n" +
"       (select rad.jy_region_name 区域,\n" +
"        decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户') 客户类型,\n" +
"               nvl(count(distinct scc.cust_id), 0) 新增数字电视客户数\n" +
"          from wxjy.jy_region_address_rel rad,\n" +
"               so1.cm_customer scc,\n" +
"               (select *\n" +
"                  from so1.ord_cust\n" +
"                   union all\n" +
"                select *\n" +
"                  from so1.ord_cust_f_2020) aa,\n" +
"               (select *\n" +
"                  from so1.ord_prod\n" +
"                  union all\n" +
"                select *\n" +
"                  from so1.ord_prod_f_2020) bb,\n" +
"                   so1.ins_prod               sip\n" +
"         where scc.cust_id = rad.cust_id\n" +
"           and aa.cust_id = scc.cust_id  and sip.cust_id=scc.cust_id\n" +
"           and aa.cust_order_id = bb.cust_order_id\n" +
"           and bb.prod_spec_id = 800200000001\n" +
"           and aa.business_id = 800001000001\n" +
"           and aa.order_state <> 10\n" +
"           and scc.own_corp_org_id = 3328\n" +
"         and scc.create_date >= trunc(sysdate) - 1\n" +
"         and scc.create_date < trunc(sysdate)\n" +
"         and exists (select 1\n" +
"          from so1.Ins_Srvpkg sis\n" +
"         where sis.prod_service_id = 1002\n" +
"           and sis.state = 1\n" +
"           and sis.os_status is null\n" +
"           and sis.prod_inst_id = sip.prod_inst_id)\n" +
"         group by rad.jy_region_name,decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户')) tp1\n" +
" where tp.区域 <> '丁蜀镇'  and tp.区域 <>'其他'\n" +
"   and tp.区域 = tp1.区域(+)\n" +
" order by tp.区域,tp1.客户类型";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 2 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\日报\\" + "日报新增电视客户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 日报缴费客户数
        public static bool baobiao5()
        {
            string sqlString =

"with t_all as\n" +
"    (select t.cust_id,t.corp_org_id,\n" +
"       MAX(CASE WHEN t.is_dtv = 1 THEN NVL(t.boder_flag,1)\n" +
"              WHEN t.is_atv = 1 THEN -1\n" +
"              END\n" +
"       ) tclx,\n" +
"       MIN(case when t.user_status = 1 then 2 when t.user_status = 2 then 1 else t.user_status end\n" +
"       ) user_status\n" +
"    from repnew.fact_ins_Prod_"+ DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " t\n" +
"    left join repnew.fact_yunweibujiqi_wx t3 on t.cust_id = t3.cust_id\n" +
"    WHERE 1=1 AND (t.is_dtv = 1 OR t.is_atv = 1)\n" +
"        and t3.cust_id is null\n" +
"    GROUP BY t.cust_id,t.corp_org_id\n" +
"    ),\n" +
" t_cancel as\n" +
"    (SELECT  t.cust_id,t.corp_org_id,\n" +
"    MAX(CASE WHEN t.is_dtv = 1 THEN NVL(t.boder_flag,1)\n" +
"              WHEN t.is_atv = 1 THEN -1\n" +
"              END\n" +
"       ) tclx,\n" +
"       MIN(case when t.user_status = 1 then 2 when t.user_status = 2 then 1 else t.user_status end\n" +
"       ) user_status\n" +
"    from repnew.fact_ins_prod_cancel t\n" +
"    where (t.is_dtv=1 OR t.is_atv=1)\n" +
"        AND NOT EXISTS (select 1 from repnew.fact_ins_Prod_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " fip WHERE t.cust_id = fip.cust_id)\n" +
"    GROUP BY t.cust_id,t.corp_org_id\n" +
"    )\n" +
"\n" +
"    select\n" +
"     tt.corp_org_id,\n" +
"     tt.区域,\n" +
"     tt.cust_type,\n" +
"       sum(tt.ktkhs) 开通客户数,\n" +
"       sum(tt.yktkhs) 预开通客户数\n" +
"       from\n" +
"(\n" +
"  select\n" +
"             t1.corp_org_id,\n" +
"             dz.jy_region_name 区域,\n" +
"             DECODE(cust.cust_type,1,'公众客户',4,'普通商业客户',7,'合同商业客户','公众客户') cust_type,\n" +
"             count(distinct case WHEN t1.user_status = 1 then t1.cust_id else null end) ktkhs, --开通客户数\n" +
"             count(distinct case when t1.user_status = 2 then t1.cust_id else null end) yktkhs --预开通客户数\n" +
"        from t_all t1\n" +
"        LEFT JOIN repnew.fact_cust_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " cust ON t1.cust_id = cust.cust_id\n" +
"        join wxjy.jy_region_address_rel dz on t1.cust_id=dz.cust_id\n" +
"       group by t1.corp_org_id,dz.jy_region_name,\n" +
"                cust.cust_type\n" +
"\n" +
"\n" +
"    UNION ALL\n" +
"    select t1.corp_org_id,\n" +
" dz.jy_region_name 区域,\n" +
"             DECODE(cust.cust_type,1,'公众客户',4,'普通商业客户',7,'合同商业客户','公众客户') cust_type,\n" +
"             0 ktkhs, --开通客户数\n" +
"             0 yktkhs --预开通客户数\n" +
"        from t_cancel t1\n" +
"        LEFT JOIN repnew.fact_cust_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " cust ON t1.cust_id = cust.cust_id\n" +
"                join  wxjy.jy_region_address_rel dz on t1.cust_id=dz.cust_id\n" +
"       group by t1.corp_org_id,dz.jy_region_name,t1.tclx, cust.cust_type\n" +
"               ) tt\n" +
"where tt.corp_org_id = 3328\n" +
" GROUP BY tt.corp_org_id,tt.区域,tt.cust_type";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\日报\\" + "日报缴费客户数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 日报新增数字电视客户数累计
        public static bool baobiao6()
        {
            string sqlString =

"select tp.区域,\n" +
"       trunc(sysdate) - 1 日期,\n" +
"        tp1.客户类型,\n" +
"       nvl(tp1.新增数字电视客户数累计, 0) 新增数字电视客户数累计\n" +
"  from (select distinct rad.jy_region_name 区域\n" +
"          from wxjy.jy_region_address_rel rad) tp,\n" +
"       (select rad.jy_region_name 区域,\n" +
"        decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户') 客户类型,\n" +
"               nvl(count(distinct scc.cust_id), 0) 新增数字电视客户数累计\n" +
"          from wxjy.jy_region_address_rel rad,\n" +
"               so1.cm_customer scc,\n" +
"               (select *\n" +
"                  from so1.ord_cust\n" +
"\n" +
"                   union all\n" +
"                select *\n" +
"                  from so1.ord_cust_f_2020) aa,\n" +
"               (select *\n" +
"                  from so1.ord_prod\n" +
"                  union all\n" +
"                select *\n" +
"                  from so1.ord_prod_f_2020) bb,\n" +
"                   so1.ins_prod               sip\n" +
"         where scc.cust_id = rad.cust_id\n" +
"           and aa.cust_id = scc.cust_id  and sip.cust_id=scc.cust_id\n" +
"           and aa.cust_order_id = bb.cust_order_id\n" +
"           and bb.prod_spec_id = 800200000001\n" +
"           and aa.business_id = 800001000001\n" +
"           and aa.order_state <> 10\n" +
"           and scc.own_corp_org_id = 3328\n" +
"         and scc.create_date >= date '" + DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString() + "-1'\n" +
"         and scc.create_date < trunc(sysdate)\n" +
"         and exists (select 1\n" +
"          from so1.Ins_Srvpkg sis\n" +
"         where sis.prod_service_id = 1002\n" +
"           and sis.state = 1\n" +
"           and sis.os_status is null\n" +
"           and sis.prod_inst_id = sip.prod_inst_id)\n" +
"         group by rad.jy_region_name,decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户')) tp1\n" +
" where tp.区域 <> '丁蜀镇'  and tp.区域 <>'其他'\n" +
"   and tp.区域 = tp1.区域(+)\n" +
" order by tp.区域,tp1.客户类型";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 2 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\日报\\" + "日报新增数字电视客户数累计--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 日报复通数字电视客户数（3个月之前停）
        public static bool baobiao7()
        {
            string sqlString =


"select rad.jy_region_name 区域,\n" +
"       trunc(sysdate) - 1 日期,\n" +
"      decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户') 客户类型,\n" +
"       count(distinct scc.cust_id) 复通数字电视客户数\n" +
"  from so1.cm_customer            scc,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_tjkh_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-3).AddDays(-1).ToString("yyyyMMdd") + "      tjkh\n" +
" where rad.cust_id = scc.cust_id\n" +
"   and scc.cust_id = tjkh.cust_id\n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and not exists (select 1\n" +
"          from repnew.fact_ins_prod_" + DateTime.Now.AddDays(-2).ToString("yyyyMMdd") + " sip2,\n" +
"               repnew.fact_ins_srvpkg_" + DateTime.Now.AddDays(-2).ToString("yyyyMMdd") + " sis2\n" +
"         where sip2.prod_inst_id = sis2.prod_inst_id\n" +
"           and sis2.prod_service_id = 1002\n" +
"           and sis2.state = '1'\n" +
"           and sis2.os_status is null\n" +
"           and sip2.cust_id = scc.cust_id)\n" +
"   and exists (select 1\n" +
"          from repnew.fact_ins_prod_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + "   sip3,\n" +
"               repnew.fact_ins_srvpkg_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " sis3\n" +
"         where sip3.prod_inst_id = sis3.prod_inst_id\n" +
"           and sis3.prod_service_id = 1002\n" +
"           and sis3.state = '1'\n" +
"           and sis3.os_status is null\n" +
"           and sip3.cust_id = scc.cust_id)\n" +
" group by rad.jy_region_name,  decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户')";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 2 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\日报\\" + "日报复通数字电视客户数（3个月之前停）--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 日报复通数字电视客户数（3个月之前停）累计
        public static bool baobiao8()
        {
            string sqlString =


"select rad.jy_region_name 区域,\n" +
"       trunc(sysdate) - 1 日期,\n" +
"      decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户') 客户类型,\n" +
"       count(distinct scc.cust_id) 复通数字电视客户数\n" +
"  from so1.cm_customer            scc,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_tjkh_"+ DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-3).AddDays(-1).ToString("yyyyMMdd") + "      tjkh\n" +
" where rad.cust_id = scc.cust_id\n" +
"   and scc.cust_id = tjkh.cust_id\n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and not exists (select 1\n" +
"          from repnew.fact_ins_prod_"+ DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sip2,\n" +
"               repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sis2\n" +
"         where sip2.prod_inst_id = sis2.prod_inst_id\n" +
"           and sis2.prod_service_id = 1002\n" +
"           and sis2.state = '1'\n" +
"           and sis2.os_status is null\n" +
"           and sip2.cust_id = scc.cust_id)\n" +
"   and exists (select 1\n" +
"          from repnew.fact_ins_prod_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + "  sip3,\n" +
"               repnew.fact_ins_srvpkg_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " sis3\n" +
"         where sip3.prod_inst_id = sis3.prod_inst_id\n" +
"           and sis3.prod_service_id = 1002\n" +
"           and sis3.state = '1'\n" +
"           and sis3.os_status is null\n" +
"           and sip3.cust_id = scc.cust_id)\n" +
" group by rad.jy_region_name,  decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户')";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 2 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\日报\\" + "日报复通数字电视客户数（3个月之前停）累计--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 日报新增宽带用户数据统计
        public static bool baobiao9()
        {
            string sqlString =
"select rad.jy_region_name 区域, count(distinct op.prod_inst_id) 新增用户数\n" +
"  from (select *\n" +
"          from so1.ord_cust\n" +
"        union all\n" +
"        select *\n" +
"          from so1.ord_cust_f_2020) oc,\n" +
"       (select *\n" +
"          from so1.ord_prod\n" +
"        union all\n" +
"        select *\n" +
"          from so1.ord_prod_f_2020) op,\n" +
"       (select *\n" +
"          from so1.ord_srvpkg\n" +
"        union all\n" +
"        select *\n" +
"          from so1.ord_srvpkg_f_2020) os,\n" +
"       so1.cm_customer scc,\n" +
"       so1.ins_address iad,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where scc.cust_id = op.cust_id\n" +
"   and scc.cust_id = iad.cust_id(+)\n" +
"   and scc.cust_id = rad.cust_id\n" +
"   and oc.CUST_ORDER_ID = op.CUST_ORDER_ID\n" +
"   and op.prod_order_id = os.prod_order_id\n" +
"   and oc.BUSINESS_ID in (800001000001, 800001000002) --普通新装,批量新装\n" +
"   and op.prod_spec_id = 800200000003 --宽带:800200000003   数字:800200000001\n" +
"   and oc.order_state <> 10 --排除新装撤单情况\n" +
"   and oc.create_date >=  trunc(sysdate)-1\n" +
"   and oc.create_date < trunc(sysdate)\n" +
"   and oc.own_corp_org_id = 3328\n" +
" group by rad.jy_region_name";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\日报\\" + "新增宽带用户数据统计--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 日报宽带缴费总用户数统计
        public static bool baobiao10()
        {
            string sqlString =
"select rad.jy_region_name 区域,\n" +
" DECODE(scc.cust_type,\n" +
"              1,\n" +
"              '住宅客户',\n" +
"              4,\n" +
"              '普通非住宅客户',\n" +
"              7,\n" +
"              '集团非住宅客户',\n" +
"              '住宅客户') 客户类型,\n" +
"       count(distinct case\n" +
"               when (t1.state = '1' and t1.is_ins = '1' and t1.os_status is null or\n" +
"                    (t1.state = '99' or t1.is_ins = '0')) and\n" +
"                    pi.name not like '%测试%' and pi.name not like '%体验%' then\n" +
"                t1.prod_inst_id\n" +
"               else\n" +
"                null\n" +
"             end) 宽带缴费用户数\n" +
"  from repnew.fact_Ins_prod_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + "   p,\n" +
"       repnew.fact_Ins_srvpkg_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " t1,\n" +
"       product.up_product_item         pi,\n" +
"       so1.cm_customer                 scc,\n" +
"       wxjy.jy_region_address_rel      rad\n" +
" where scc.cust_id = rad.cust_id\n" +
"   and p.cust_id = scc.cust_id\n" +
"   and p.prod_inst_id = t1.prod_inst_id\n" +
"   and t1.prod_service_id = 1004\n" +
"   and p.corp_org_id = 3328\n" +
"   and t1.srvpkg_id = pi.product_item_id\n" +
" group by rad.jy_region_name, DECODE(scc.cust_type,\n" +
"              1,\n" +
"              '住宅客户',\n" +
"              4,\n" +
"              '普通非住宅客户',\n" +
"              7,\n" +
"              '集团非住宅客户',\n" +
"              '住宅客户')";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\日报\\" + "宽带缴费总用户数统计--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 日报学霸宝盒总订购量客户明细数据
        public static bool baobiao11()
        {
            string sqlString =
"select scc.cust_code 客户证号,\n" +
"       scc.cust_name 客户姓名,\n" +
"          decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户') 客户类型,\n" +
"       ipr.res_equ_no 机顶盒号码,\n" +
"       sip.sub_bill_id 智能卡,\n" +
"       upd.name 产品名称,\n" +
"       sis.create_date 订购时间,\n" +
"       sor.organize_name 营业厅,\n" +
"       sta.staff_name 操作员,\n" +
"       con.cont_phone1 || ',' || con.cont_phone2 联系方式,\n" +
"       con.cont_mobile1 || ',' || con.cont_mobile2 移动电话,\n" +
"       rad.jy_region_name 区域,\n" +
"       iad.std_addr_name || iad.door_name 地址\n" +
"  from so1.cm_customer            scc,\n" +
"       so1.ins_prod               sip,\n" +
"       so1.ins_srvpkg             sis,\n" +
"       product.up_product_item    upd,\n" +
"       so1.ins_prod_res           ipr,\n" +
"       sec.sec_operator           sop,\n" +
"       sec.sec_staff              sta,\n" +
"       so1.ins_address            iad,\n" +
"       so1.cm_cust_contact_info   con,\n" +
"       sec.sec_organize           sor,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where scc.cust_id = sip.cust_id\n" +
"   and sip.prod_inst_id = sis.prod_inst_id\n" +
"   and sis.srvpkg_id = upd.product_item_id\n" +
"   and ipr.prod_inst_id = sip.prod_inst_id\n" +
"   and sis.op_id = sop.operator_id\n" +
"   and sta.staff_id = sop.staff_id\n" +
"   and scc.cust_id = iad.cust_id\n" +
"   and scc.cust_id = con.cust_id\n" +
"   and sta.organize_id = sor.organize_id\n" +
"   and scc.cust_id = rad.cust_id\n" +
"   and (upd.name like '%学霸宝盒OTT%')\n" +
"   and sis.state = 1\n" +
"   and sis.os_status is null\n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and sis.create_date < trunc(sysdate)\n" +
"   and exists (select 1\n" +
"          from res.res_code_definition rcd\n" +
"         where rcd.res_type = 2\n" +
"           and ipr.res_code = rcd.res_code)\n" +
" order by sis.create_date, scc.cust_code";

            int[] columntxt = { 1, 4, 5, 10, 11 }; //哪些列是文本格式
            int[] columndate = { 7 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\日报\\" + "学霸宝盒总订购客户明细数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 日报高清互动终端新增统计
        public static bool baobiao12()
        {
            string sqlString =
"select rad.jy_region_name 区域,\n" +
"     decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户') 客户类型,\n" +
"       count(distinct ipr.res_equ_no) 新增高清机顶盒数量\n" +
"\n" +
"  from so1.cm_customer            scc,\n" +
"       so1.ins_prod               sip,\n" +
"       so1.ins_prod_res           ipr,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where scc.cust_id = rad.cust_id\n" +
"   and scc.cust_id = sip.cust_id\n" +
"   and sip.prod_inst_id = ipr.prod_inst_id\n" +
"   and sip.sub_bill_id is not null\n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and ipr.create_date >= trunc(sysdate)-1\n" +
"   and ipr.create_date < trunc(sysdate)\n" +
"   and exists\n" +
" (select 1\n" +
"          from res.res_terminal rrt, res.res_code_definition rcd\n" +
"         where rrt.res_code = rcd.res_code\n" +
"             and rcd.res_name in ('银河高清交互型HDC691033(江阴)',\n" +
"                        '4K超高清型II型融合型（EOC）',\n" +
"                        '银河智能高清交互型HDC6910798(江阴)',\n" +
"                        '银河4K交互型HDC691090',\n" +
"                        '银河4K超高清II型融合型HDC6910B1(江阴)',\n" +
"                        '银河高清基本型HDC6910(江阴)')\n" +
"           and ipr.res_equ_no = rrt.serial_no)\n" +
" group by rad.jy_region_name,     decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户')";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\日报\\" + "高清互动终端新增统计--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 日报总高清互动终端缴费总数统计
        public static bool baobiao13()
        {
            string sqlString =
            "select trim(case\n" +
            "              when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
            "                   scc.std_addr_name not like '%测试%' then\n" +
            "               substr(scc.std_addr_name, 13, 3)\n" +
            "              when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
            "                   scc.std_addr_name not like '%测试%' then\n" +
            "               substr(scc.std_addr_name, 6, 3)\n" +
            "              when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
            "               '澄江'\n" +
            "              when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
            "               '华士'\n" +
            "              when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
            "               '南闸'\n" +
            "              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
            "               '申港'\n" +
            "              when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
            "               '祝塘'\n" +
            "              when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
            "               '其他'\n" +
            "              else\n" +
            "               substr(scc.std_addr_name, 1, 3)\n" +
            "            end) 区域,\n" +
            "       DECODE(scc.cust_type,\n" +
            "              1,\n" +
            "              '住宅客户',\n" +
            "              4,\n" +
            "              '普通非住宅客户',\n" +
            "              7,\n" +
            "              '集团非住宅客户',\n" +
            "              '住宅客户') 客户类型,\n" +
            "       count(distinct sip.bill_id) 高清互动缴费终端总数\n" +
            "  from repnew.fact_ins_prod_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + "  sip,\n" +
            "       repnew.fact_cust_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + "     scc\n" +
            " where sip.cust_id = scc.cust_id\n" +
            "   and scc.own_corp_org_id = 3328\n" +
            "   and exists (select 1\n" +
            "          from repnew.fact_ins_srvpkg_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " sis1\n" +
            "         where sis1.prod_service_id = 1002\n" +
            "           and sis1.state = 1\n" +
            "           and sis1.os_status is null\n" +
            "           and sis1.prod_inst_id = sip.prod_inst_id) --有正常基本包\n" +
            "   and exists\n" +
            " (select 1\n" +
            "          from res.res_terminal rrt, res.res_code_definition rcd\n" +
            "         where rrt.res_code = rcd.res_code\n" +
            "           and (rcd.res_name like '%高清%' or rcd.res_name like '%4K%')\n" +
            "           and sip.bill_id = rrt.serial_no) --有高清或者4K机顶盒\n" +
            "and not exists\n" +
            " (select 1\n" +
            "          from res.res_terminal rrt, res.res_code_definition rcd\n" +
            "         where rrt.res_code = rcd.res_code\n" +
            "           and rcd.res_name in ('九州高清基本型DVC7058(江阴)','大亚高清交互型DC5000(江阴)','天柏高清基本型HMC0201BDH(江阴)','创维高清互动机顶盒_常熟','创维高清一体机(仪征)','九州高清交互型DVC7058EOC(江阴)','同洲高清交互型N8606(江阴)','同洲高清交互型5120de(江阴)')\n" +
            "           and sip.bill_id = rrt.serial_no) --排除干扰项高清类型\n" +
            "   and exists (select 1\n" +
            "          from repnew.fact_ins_srvpkg_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " sis,\n" +
            "               wxjy.rep_jy_itv_product         jip\n" +
            "         where sis.srvpkg_id = jip.product_item_id\n" +
            "           and sis.prod_inst_id = sip.prod_inst_id) --有互动产品\n" +
            " group by trim(case\n" +
            "                 when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
            "                      scc.std_addr_name not like '%测试%' then\n" +
            "                  substr(scc.std_addr_name, 13, 3)\n" +
            "                 when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
            "                      scc.std_addr_name not like '%测试%' then\n" +
            "                  substr(scc.std_addr_name, 6, 3)\n" +
            "                 when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
            "                  '澄江'\n" +
            "                 when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
            "                  '华士'\n" +
            "                 when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
            "                  '南闸'\n" +
            "              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
            "               '申港'\n" +
            "                 when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
            "                  '祝塘'\n" +
            "                 when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
            "                  '其他'\n" +
            "                 else\n" +
            "                  substr(scc.std_addr_name, 1, 3)\n" +
            "               end),\n" +
            "          DECODE(scc.cust_type,\n" +
            "                 1,\n" +
            "                 '住宅客户',\n" +
            "                 4,\n" +
            "                 '普通非住宅客户',\n" +
            "                 7,\n" +
            "                 '集团非住宅客户',\n" +
            "                 '住宅客户')";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\日报\\" + "高清互动终端缴费总数统计--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion








        #region 周报产品订购
        public static bool zbaobiao1()
        {
            string sqlString =
"select scc.cust_code 客户证号,\n" +
"       scc.cust_name 客户姓名,\n" +
"                  decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户') 客户类型,\n" +
"       ipr.res_equ_no 机顶盒号码,\n" +
"       sip.sub_bill_id 智能卡,\n" +
"       upd.name 产品名称,\n" +
"       sis.create_date 订购时间,\n" +
"       sor.organize_name 营业厅,\n" +
"       sta.staff_name 操作员,\n" +
"       con.cont_phone1 || ',' || con.cont_phone2 联系方式,\n" +
"       con.cont_mobile1 || ',' || con.cont_mobile2 移动电话,\n" +
"       rad.jy_region_name 区域,\n" +
"       iad.std_addr_name || iad.door_name 地址\n" +
"  from so1.cm_customer            scc,\n" +
"       so1.ins_prod               sip,\n" +
"       so1.ins_srvpkg             sis,\n" +
"       product.up_product_item    upd,\n" +
"       so1.ins_prod_res           ipr,\n" +
"       sec.sec_operator           sop,\n" +
"       sec.sec_staff              sta,\n" +
"       so1.ins_address            iad,\n" +
"       so1.cm_cust_contact_info   con,\n" +
"       sec.sec_organize           sor,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where scc.cust_id = sip.cust_id\n" +
"   and sip.prod_inst_id = sis.prod_inst_id\n" +
"   and sis.srvpkg_id = upd.product_item_id\n" +
"   and ipr.prod_inst_id = sip.prod_inst_id\n" +
"   and sis.op_id = sop.operator_id\n" +
"   and sta.staff_id = sop.staff_id\n" +
"   and scc.cust_id = iad.cust_id\n" +
"   and scc.cust_id = con.cust_id\n" +
"   and sta.organize_id = sor.organize_id\n" +
"   and scc.cust_id = rad.cust_id\n" +
"   and (upd.name like '%芒果%' or upd.name like '%七彩童年%' or upd.name like '%电竞频道%')\n" +
"   and sis.state = 1\n" +
"   and sis.os_status is null\n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and sis.create_date >= trunc(sysdate)-7\n" +
"   and sis.create_date < trunc(sysdate)\n" +
"   and exists (select 1\n" +
"          from res.res_code_definition rcd\n" +
"         where rcd.res_type = 2\n" +
"           and ipr.res_code = rcd.res_code)";

            int[] columntxt = { 1,4,5,10,11 }; //哪些列是文本格式
            int[] columndate = { 7 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\周报\\" + "周报产品订购--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 周报产品退订
        public static bool zbaobiao2()
        {
            string sqlString =
"select distinct scc.cust_code      客户证号,\n" +
"                scc.cust_name      客户姓名,\n" +
"                prod.bill_id       机顶盒号码,\n" +
"                upd.name           产品名称,\n" +
"                inof.create_date   订购时间,\n" +
"                spkg.create_date   退订时间,\n" +
"                sor.organize_name  营业厅,\n" +
"                sta.staff_name     操作员,\n" +
"                rad.jy_region_name 区域,\n" +
"                iad.std_addr_name  地址\n" +
"  from so1.cm_customer scc,\n" +
"       (select *\n" +
"          from so1.ord_cust\n" +
"        union all\n" +
"        select *\n" +
"          from so1.ord_cust_f_2020) cust,\n" +
"       (select *\n" +
"          from so1.ord_prod\n" +
"        union all\n" +
"        select *\n" +
"          from so1.ord_prod_f_2020) prod,\n" +
"       (select *\n" +
"          from so1.ord_srvpkg\n" +
"        union all\n" +
"        select *\n" +
"          from so1.ord_srvpkg_f_2020) spkg,\n" +
"       (select *\n" +
"          from so1.ord_offer\n" +
"        union all\n" +
"        select *\n" +
"          from so1.ord_offer_f_2020) offe,\n" +
"       (select * from so1.h_ins_offer_2020) inof,\n" +
"       product.up_product_item upd,\n" +
"       so1.ins_address iad,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       sec.sec_operator sop,\n" +
"       sec.sec_staff sta,\n" +
"       sec.sec_organize sor\n" +
" where scc.cust_id = cust.cust_id\n" +
"   and scc.cust_id = iad.cust_id\n" +
"   and scc.cust_id = rad.cust_id\n" +
"   and cust.cust_order_id = prod.cust_order_id\n" +
"   and prod.prod_order_id = spkg.prod_order_id\n" +
"   and offe.offer_order_id = spkg.offer_order_id\n" +
"   and offe.offer_inst_id = inof.offer_inst_id\n" +
"   and spkg.srvpkg_id = upd.product_item_id\n" +
"   and cust.op_id = sop.operator_id(+)\n" +
"   and sop.staff_id = sta.staff_id(+)\n" +
"   and sta.organize_id = sor.organize_id(+)\n" +
"   and (upd.name like '%芒果%' or upd.name like '%七彩童年%' or upd.name like '%电竞频道%')\n" +
"   and cust.business_id = 800001000026\n" +
"   and spkg.state = 3\n" +
"   and spkg.create_date >= trunc(sysdate) - 7\n" +
"   and spkg.create_date < trunc(sysdate)\n" +
"   and scc.own_corp_org_id = 3328\n" +
" order by spkg.create_date";
            int[] columntxt = { 1, 3 }; //哪些列是文本格式
            int[] columndate = { 5,6 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\周报\\" + "周报产品退订--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 周报炫力芒果七彩总订购量客户明细数据
        public static bool zbaobiao3()
        {
            string sqlString =

"select scc.cust_code 客户证号,\n" +
"       scc.cust_name 客户姓名,\n" +
"          decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户') 客户类型,\n" +
"       ipr.res_equ_no 机顶盒号码,\n" +
"       sip.sub_bill_id 智能卡,\n" +
"       upd.name 产品名称,\n" +
"       sis.create_date 订购时间,\n" +
"       sor.organize_name 营业厅,\n" +
"       sta.staff_name 操作员,\n" +
"       con.cont_phone1 || ',' || con.cont_phone2 联系方式,\n" +
"       con.cont_mobile1 || ',' || con.cont_mobile2 移动电话,\n" +
"       rad.jy_region_name 区域,\n" +
"       iad.std_addr_name || iad.door_name 地址\n" +
"  from so1.cm_customer            scc,\n" +
"       so1.ins_prod               sip,\n" +
"       so1.ins_srvpkg             sis,\n" +
"       product.up_product_item    upd,\n" +
"       so1.ins_prod_res           ipr,\n" +
"       sec.sec_operator           sop,\n" +
"       sec.sec_staff              sta,\n" +
"       so1.ins_address            iad,\n" +
"       so1.cm_cust_contact_info   con,\n" +
"       sec.sec_organize           sor,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where scc.cust_id = sip.cust_id\n" +
"   and sip.prod_inst_id = sis.prod_inst_id\n" +
"   and sis.srvpkg_id = upd.product_item_id\n" +
"   and ipr.prod_inst_id = sip.prod_inst_id\n" +
"   and sis.op_id = sop.operator_id\n" +
"   and sta.staff_id = sop.staff_id\n" +
"   and scc.cust_id = iad.cust_id\n" +
"   and scc.cust_id = con.cust_id\n" +
"   and sta.organize_id = sor.organize_id\n" +
"   and scc.cust_id = rad.cust_id\n" +
"   and (upd.name like '%炫力动漫%' or upd.name like '%芒果%' or upd.name like '%七彩童年%')\n" +
"   and sis.state = 1\n" +
"   and sis.os_status is null\n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and sis.create_date < trunc(sysdate)\n" +
"   and exists (select 1\n" +
"          from res.res_code_definition rcd\n" +
"         where rcd.res_type = 2\n" +
"           and ipr.res_code = rcd.res_code)\n" +
" order by sis.create_date, scc.cust_code";

            int[] columntxt = { 1, 4,5,10,11 }; //哪些列是文本格式
            int[] columndate = { 7 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\周报\\" + "周报炫力芒果七彩总订购量客户明细数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 周报电视交警点播套餐明细全量
        public static bool zbaobiao4()
        {
            string sqlString =
"select distinct scc.cust_code 客户证号,\n" +
"       scc.cust_name 客户姓名,\n" +
"       DECODE(scc.cust_type,1,'住宅客户',4, '普通非住宅客户',7,'集团客户','住宅客户') 客户类型,\n" +
"       upd.name 频道名称,\n" +
"         sip.bill_id 机顶盒号码,\n" +
"        decode(sis.os_status,\n" +
"                                  1,\n" +
"                                  '欠费停',\n" +
"                                  3,\n" +
"                                  '暂停',\n" +
"                                  4,\n" +
"                                  '管理停机',\n" +
"                                  '正常') 状态,\n" +
"       con.cont_phone1 || ',' || con.cont_phone2 联系方式,\n" +
"                con.cont_mobile1 || ',' || con.cont_mobile2 移动电话,\n" +
"                rad.jy_region_name 区域,\n" +
"                iad.std_addr_name || iad.door_name 地址\n" +
"  from so1.cm_customer         scc,\n" +
"       so1.ins_prod            sip,\n" +
"       so1.ins_srvpkg          sis,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"        so1.ins_address            iad,\n" +
"       so1.cm_cust_contact_info   con,\n" +
"       product.up_product_item upd\n" +
" where scc.cust_id = sip.cust_id\n" +
"   and sip.prod_inst_id = sis.prod_inst_id\n" +
"     and scc.cust_id = iad.cust_id(+)\n" +
"   and scc.cust_id = con.cust_id(+)\n" +
"   and scc.cust_id = rad.cust_id\n" +
"   and sis.srvpkg_id = upd.product_item_id\n" +
"   and (upd.name like '%电视交警_点播首月300元-次年25元/月_江阴%')\n" +
"      and sis.create_date >= trunc(sysdate)-7\n" +
"   and sis.create_date < trunc(sysdate)\n" +
"   and scc.own_corp_org_id = 3328";

            int[] columntxt = { 1,  5, 7, 8 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\周报\\" + "周报电视交警点播套餐明细全量--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 周报电视交警宽带套餐明细全量
        public static bool zbaobiao5()
        {
            string sqlString =
"select distinct scc.cust_code 客户证号,\n" +
"       scc.cust_name 客户姓名,\n" +
"       DECODE(scc.cust_type,1,'住宅客户',4, '普通非住宅客户',7,'集团客户','住宅客户') 客户类型,\n" +
"       upd.name 频道名称,\n" +
"         sip.bill_id 机顶盒号码,\n" +
"        decode(sis.os_status,\n" +
"                                  1,\n" +
"                                  '欠费停',\n" +
"                                  3,\n" +
"                                  '暂停',\n" +
"                                  4,\n" +
"                                  '管理停机',\n" +
"                                  '正常') 状态,\n" +
"       con.cont_phone1 || ',' || con.cont_phone2 联系方式,\n" +
"                con.cont_mobile1 || ',' || con.cont_mobile2 移动电话,\n" +
"                rad.jy_region_name 区域,\n" +
"                iad.std_addr_name || iad.door_name 地址\n" +
"  from so1.cm_customer         scc,\n" +
"       so1.ins_prod            sip,\n" +
"       so1.ins_srvpkg          sis,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"        so1.ins_address            iad,\n" +
"       so1.cm_cust_contact_info   con,\n" +
"       product.up_product_item upd\n" +
" where scc.cust_id = sip.cust_id\n" +
"   and sip.prod_inst_id = sis.prod_inst_id\n" +
"     and scc.cust_id = iad.cust_id(+)\n" +
"   and scc.cust_id = con.cust_id(+)\n" +
"   and scc.cust_id = rad.cust_id\n" +
"   and sis.srvpkg_id = upd.product_item_id\n" +
"   and (upd.name like '%电视交警50M宽带_首月60元-次年5元/月_江阴%')\n" +
"      and sis.create_date >= trunc(sysdate)-7\n" +
"   and sis.create_date < trunc(sysdate)\n" +
"   and scc.own_corp_org_id = 3328";

            int[] columntxt = { 1, 5, 7, 8 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\周报\\" + "周报电视交警宽带套餐明细全量--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 周报畅看系列明细
        public static bool zbaobiao6()
        {
            string sqlString =

"select scc.cust_code 客户证号,\n" +
"       scc.cust_name 客户姓名,\n" +
"          decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户') 客户类型,\n" +
"       ipr.res_equ_no 机顶盒号码,\n" +
"       sip.sub_bill_id 智能卡,\n" +
"       upd.name 产品名称,\n" +
"       sis.create_date 订购时间,\n" +
"       sor.organize_name 营业厅,\n" +
"       sta.staff_name 操作员,\n" +
"       con.cont_phone1 || ',' || con.cont_phone2 联系方式,\n" +
"       con.cont_mobile1 || ',' || con.cont_mobile2 移动电话,\n" +
"       rad.jy_region_name 区域,\n" +
"       iad.std_addr_name || iad.door_name 地址\n" +
"  from so1.cm_customer            scc,\n" +
"       so1.ins_prod               sip,\n" +
"       so1.ins_srvpkg             sis,\n" +
"       product.up_product_item    upd,\n" +
"       so1.ins_prod_res           ipr,\n" +
"       sec.sec_operator           sop,\n" +
"       sec.sec_staff              sta,\n" +
"       so1.ins_address            iad,\n" +
"       so1.cm_cust_contact_info   con,\n" +
"       sec.sec_organize           sor,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where scc.cust_id = sip.cust_id\n" +
"   and sip.prod_inst_id = sis.prod_inst_id\n" +
"   and sis.srvpkg_id = upd.product_item_id\n" +
"   and ipr.prod_inst_id = sip.prod_inst_id\n" +
"   and sis.op_id = sop.operator_id\n" +
"   and sta.staff_id = sop.staff_id\n" +
"   and scc.cust_id = iad.cust_id\n" +
"   and scc.cust_id = con.cust_id\n" +
"   and sta.organize_id = sor.organize_id\n" +
"   and scc.cust_id = rad.cust_id\n" +
"   and upd.name like '%畅看%'\n" +
"   and sis.state = 1\n" +
"   and sis.os_status is null\n" +
"   and scc.own_corp_org_id = 3328\n" +
"  and sis.create_date >= trunc(sysdate)-7\n" +
"   and sis.create_date < trunc(sysdate)\n" +
"   and exists (select 1\n" +
"          from res.res_code_definition rcd\n" +
"         where rcd.res_type = 2\n" +
"           and ipr.res_code = rcd.res_code)\n" +
" order by sis.create_date, scc.cust_code";

            int[] columntxt = { 1,4, 5, 10, 11 }; //哪些列是文本格式
            int[] columndate = { 7 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\周报\\" + "周报畅看系列明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 周报看视界购机促销标和清升4K套餐订购明细
        public static bool zbaobiao7()
        {
            string sqlString =

"select scc.cust_code 客户证号,\n" +
"       scc.cust_name 客户姓名,\n" +
"          decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户') 客户类型,\n" +
"       ipr.res_equ_no 机顶盒号码,\n" +
"sof.order_name 套餐名称,\n" +
"sof.create_date 套餐订购时间,\n" +
"       sor.organize_name 营业厅,\n" +
"       sta.staff_name 操作员,\n" +
"       rad.jy_region_name 区域\n" +
"  from so1.cm_customer          scc,\n" +
"       so1.ins_prod             sip,\n" +
"       so1.ins_off_ins_prod_rel rel,\n" +
"       so1.ins_offer            sof,\n" +
"       sec.sec_staff            sta,\n" +
"       sec.sec_operator         sop,\n" +
"       sec.sec_organize         sor,\n" +
"       so1.ins_prod_res           ipr,\n" +
"       res.res_terminal           rrt,\n" +
"       res.res_code_definition    rcd,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where scc.cust_id = sip.cust_id\n" +
"   and sip.prod_inst_id = rel.prod_inst_id\n" +
"   and rel.offer_inst_id = sof.offer_inst_id\n" +
"   and sof.op_id = sop.operator_id\n" +
"   and sop.staff_id = sta.staff_id\n" +
"   and sof.org_id = sor.organize_id\n" +
"   and sip.prod_inst_id = ipr.prod_inst_id\n" +
"   and rad.cust_id = scc.cust_id\n" +
"   and ipr.res_equ_no = rrt.serial_no\n" +
"   and rrt.res_code = rcd.res_code\n" +
"   and rcd.res_type = 2\n" +
"   and (sof.order_name like '%看视界购机促销套餐%' or sof.order_name like '%标升4K%')\n" +
"   and sof.create_date >= trunc(sysdate)-7\n" +
"   and sof.create_date < trunc(sysdate)\n" +
"   and scc.own_corp_org_id = 3328";

            int[] columntxt = { 1, 4 }; //哪些列是文本格式
            int[] columndate = { 6 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\周报\\" + "周报看视界购机促销标和清升4K套餐订购明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 周报开具288或399或499机顶盒发票的客户明细
        public static bool zbaobiao8()
        {
            string sqlString =
"select distinct a.invoice_list_no 发票代码,\n" +
"               a.invoice_num 发票号码,\n" +
"               a.invoice_amount / 100 发票金额,\n" +
"               a.image 科目,\n" +
"               decode(a.invoice_state, 1, '正常', 3, '冲红', 4, '被冲红', '') 发票类型,\n" +
"               a.print_date 开票时间,\n" +
"               a.remark 备注,\n" +
"               c.cust_code 客户证号,\n" +
"               c.cust_name 客户姓名,\n" +
"               c.create_date  客户新装时间,\n" +
"               decode(c.cust_type,\n" +
"                      '1',\n" +
"                      '公众客户',\n" +
"                      '4',\n" +
"                      '普通商业客户',\n" +
"                      '7',\n" +
"                      '合同商业客户') 客户类型,\n" +
"               b.staff_name 开票人,\n" +
"               e.organize_name 营业厅\n" +
"          from so1.ord_invoice_2020    a,\n" +
"               so1.cm_customer         c,\n" +
"               sec.sec_staff           b,\n" +
"               sec.sec_operator        d,\n" +
"               sec.sec_organize        e,\n" +
"               so1.ins_prod            sip,\n" +
"               so1.ins_prod_res        ipr,\n" +
"               res.res_terminal        rrt,\n" +
"               res.res_code_definition rcd\n" +
"         where ipr.prod_inst_id = sip.prod_inst_id\n" +
"           and c.cust_id = sip.cust_id\n" +
"           and rrt.res_code = rcd.res_code\n" +
"           and sip.bill_id = rrt.serial_no\n" +
"           and a.cust_id = c.cust_id\n" +
"           and c.own_corp_org_id = 3328\n" +
"           and a.op_id = d.operator_id\n" +
"           and d.staff_id = b.staff_id\n" +
"           and a.org_id = e.organize_id\n" +
"           and (a.invoice_amount = 28800 or a.invoice_amount = 39900 or a.invoice_amount = 49900)\n" +
"             and a.print_date  >= trunc(sysdate)-7\n" +
"   and a.print_date  < trunc(sysdate)";

            int[] columntxt = { 1, 2,8 }; //哪些列是文本格式
            int[] columndate = { 6,10 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\周报\\" + "周报开具288或399或499机顶盒发票的客户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 周报新增高清机顶盒数量累计
        public static bool zbaobiao9()
        {
            string sqlString =

"select rad.jy_region_name 区域,\n" +
"     decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户') 客户类型,\n" +
"       count(distinct ipr.res_equ_no) 新增高清机顶盒数量\n" +
"\n" +
"  from so1.cm_customer            scc,\n" +
"       so1.ins_prod               sip,\n" +
"       so1.ins_prod_res           ipr,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where scc.cust_id = rad.cust_id\n" +
"   and scc.cust_id = sip.cust_id\n" +
"   and sip.prod_inst_id = ipr.prod_inst_id\n" +
"   and sip.sub_bill_id is not null\n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and ipr.create_date >= trunc(sysdate)-7\n" +
"   and ipr.create_date < trunc(sysdate)\n" +
"   and exists\n" +
" (select 1\n" +
"          from res.res_terminal rrt, res.res_code_definition rcd\n" +
"         where rrt.res_code = rcd.res_code\n" +
"           and (rcd.res_name like '%高清%' or rcd.res_name like '%4K%')\n" +
"           and ipr.res_equ_no = rrt.serial_no)\n" +
" group by rad.jy_region_name,     decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户')";


            int[] columntxt = {0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\周报\\" + "周报新增高清机顶盒数量累计--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 周报2020TCL电视机明细
        public static bool zbaobiao10()
        {
            string sqlString =

"select scc.cust_code 客户证号,\n" +
"       scc.cust_name 客户姓名,\n" +
"       ipr.res_equ_no 机顶盒号码,\n" +
"       sip.sub_bill_id 智能卡,\n" +
"       upd.name 产品名称,\n" +
"       sis.create_date 订购时间,\n" +
"       sor.organize_name 营业厅,\n" +
"       sta.staff_name 操作员,\n" +
"       con.cont_phone1 || ',' || con.cont_phone2 联系方式,\n" +
"       con.cont_mobile1 || ',' || con.cont_mobile2 移动电话,\n" +
"       rad.jy_region_name 区域,\n" +
"       iad.std_addr_name || iad.door_name 地址\n" +
"  from so1.cm_customer            scc,\n" +
"       so1.ins_prod               sip,\n" +
"       so1.ins_srvpkg             sis,\n" +
"       product.up_product_item    upd,\n" +
"       so1.ins_prod_res           ipr,\n" +
"       sec.sec_operator           sop,\n" +
"       sec.sec_staff              sta,\n" +
"       so1.ins_address            iad,\n" +
"       so1.cm_cust_contact_info   con,\n" +
"       sec.sec_organize           sor,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where scc.cust_id = sip.cust_id\n" +
"   and sip.prod_inst_id = sis.prod_inst_id\n" +
"   and sis.srvpkg_id = upd.product_item_id\n" +
"   and ipr.prod_inst_id = sip.prod_inst_id\n" +
"   and sis.op_id = sop.operator_id\n" +
"   and sta.staff_id = sop.staff_id\n" +
"   and scc.cust_id = iad.cust_id\n" +
"   and scc.cust_id = con.cust_id\n" +
"   and sta.organize_id = sor.organize_id\n" +
"   and scc.cust_id = rad.cust_id\n" +
"   and (upd.name like '%2020TCL电视机%')\n" +
"   and sis.state = 1\n" +
"   and sis.os_status is null\n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and sis.create_date >= trunc(sysdate)-7\n" +
"   and sis.create_date < trunc(sysdate)\n" +
"   and exists (select 1\n" +
"          from res.res_code_definition rcd\n" +
"         where rcd.res_type = 2\n" +
"           and ipr.res_code = rcd.res_code)\n" +
" order by sis.create_date, scc.cust_code";
            int[] columntxt = { 1,3,4,10 }; //哪些列是文本格式
            int[] columndate = { 6 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\周报\\" + "周报2020TCL电视机明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 周报TCL各型号购买发票明细
        public static bool zbaobiao12()
        {
            string sqlString =


"select distinct a.invoice_list_no 发票代码,\n" +
"               a.invoice_num 发票号码,\n" +
"               a.invoice_amount / 100 发票金额,\n" +
"               a.image 科目,\n" +
"               decode(a.invoice_state, 1, '正常', 3, '冲红', 4, '被冲红', '') 发票类型,\n" +
"               a.print_date 开票时间,\n" +
"               a.remark 备注,\n" +
"               c.cust_code 客户证号,\n" +
"               c.cust_name 客户姓名,\n" +
"               c.create_date  客户新装时间,\n" +
"               decode(c.cust_type,\n" +
"                      '1',\n" +
"                      '公众客户',\n" +
"                      '4',\n" +
"                      '普通商业客户',\n" +
"                      '7',\n" +
"                      '合同商业客户') 客户类型,\n" +
"               b.staff_name 开票人,\n" +
"               e.organize_name 营业厅\n" +
"          from so1.ord_invoice_2020    a,\n" +
"               so1.cm_customer         c,\n" +
"               sec.sec_staff           b,\n" +
"               sec.sec_operator        d,\n" +
"               sec.sec_organize        e,\n" +
"               so1.ins_prod            sip,\n" +
"               so1.ins_prod_res        ipr,\n" +
"               res.res_terminal        rrt,\n" +
"               res.res_code_definition rcd\n" +
"         where ipr.prod_inst_id = sip.prod_inst_id\n" +
"           and c.cust_id = sip.cust_id\n" +
"           and rrt.res_code = rcd.res_code\n" +
"           and sip.bill_id = rrt.serial_no\n" +
"           and a.cust_id = c.cust_id\n" +
"           and c.own_corp_org_id = 3328\n" +
"           and a.op_id = d.operator_id\n" +
"           and d.staff_id = b.staff_id\n" +
"           and a.org_id = e.organize_id\n" +
"           and (a.image like '%电视机%' or a.image like '%TCL%' )\n" +
"   and a.print_date  >= trunc(sysdate)-7\n" +
"   and a.print_date  < trunc(sysdate)";




            int[] columntxt = { 1, 2, 8 }; //哪些列是文本格式
            int[] columndate = { 6,10 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\周报\\" + "周报TCL各型号购买发票明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 周报看视界系列明细
        public static bool zbaobiao13()
        {
            string sqlString =
"select scc.cust_code 客户证号,\n" +
"       scc.cust_name 客户姓名,\n" +
"          decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户') 客户类型,\n" +
"       ipr.res_equ_no 机顶盒号码,\n" +
"       sip.sub_bill_id 智能卡,\n" +
"       upd.name 产品名称,\n" +
"       sis.create_date 订购时间,\n" +
"       sor.organize_name 营业厅,\n" +
"       sta.staff_name 操作员,\n" +
"       con.cont_phone1 || ',' || con.cont_phone2 联系方式,\n" +
"       con.cont_mobile1 || ',' || con.cont_mobile2 移动电话,\n" +
"       rad.jy_region_name 区域,\n" +
"       iad.std_addr_name || iad.door_name 地址\n" +
"  from so1.cm_customer            scc,\n" +
"       so1.ins_prod               sip,\n" +
"       so1.ins_srvpkg             sis,\n" +
"       product.up_product_item    upd,\n" +
"       so1.ins_prod_res           ipr,\n" +
"       sec.sec_operator           sop,\n" +
"       sec.sec_staff              sta,\n" +
"       so1.ins_address            iad,\n" +
"       so1.cm_cust_contact_info   con,\n" +
"       sec.sec_organize           sor,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where scc.cust_id = sip.cust_id\n" +
"   and sip.prod_inst_id = sis.prod_inst_id\n" +
"   and sis.srvpkg_id = upd.product_item_id\n" +
"   and ipr.prod_inst_id = sip.prod_inst_id\n" +
"   and sis.op_id = sop.operator_id\n" +
"   and sta.staff_id = sop.staff_id\n" +
"   and scc.cust_id = iad.cust_id\n" +
"   and scc.cust_id = con.cust_id\n" +
"   and sta.organize_id = sor.organize_id\n" +
"   and scc.cust_id = rad.cust_id\n" +
"   and (upd.name like '%看视界%')\n" +
"   and sis.state = 1\n" +
"   and sis.os_status is null\n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and sis.create_date >= trunc(sysdate)-7\n" +
"   and sis.create_date < trunc(sysdate)\n" +
"   and exists (select 1\n" +
"          from res.res_code_definition rcd\n" +
"         where rcd.res_type = 2\n" +
"           and ipr.res_code = rcd.res_code)\n" +
" order by sis.create_date, scc.cust_code";

            int[] columntxt = { 1, 4, 5,10,11 }; //哪些列是文本格式
            int[] columndate = { 7 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\周报\\" + "周报看视界系列明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 周报新增商业客户数
        public static bool zbaobiao14()
        {
            string sqlString =
"select tp.区域,\n" +
"       trunc(sysdate) - 1 日期,\n" +
"        tp1.客户类型,\n" +
"       nvl(tp1.新增数字电视客户数, 0) 新增商业客户数\n" +
"  from (select distinct rad.jy_region_name 区域\n" +
"          from wxjy.jy_region_address_rel rad) tp,\n" +
"       (select rad.jy_region_name 区域,\n" +
"        decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户') 客户类型,\n" +
"               nvl(count(distinct scc.cust_id), 0) 新增数字电视客户数\n" +
"          from wxjy.jy_region_address_rel rad,\n" +
"               so1.cm_customer scc,\n" +
"               (select *\n" +
"                  from so1.ord_cust\n" +
"                   union all\n" +
"                select *\n" +
"                  from so1.ord_cust_f_2020) aa,\n" +
"               (select *\n" +
"                  from so1.ord_prod\n" +
"                  union all\n" +
"                select *\n" +
"                  from so1.ord_prod_f_2020) bb,\n" +
"                   so1.ins_prod               sip\n" +
"         where scc.cust_id = rad.cust_id\n" +
"           and aa.cust_id = scc.cust_id  and sip.cust_id=scc.cust_id\n" +
"           and aa.cust_order_id = bb.cust_order_id\n" +
"           and bb.prod_spec_id = 800200000001\n" +
"           and aa.business_id = 800001000001\n" +
"           and aa.order_state <> 10\n" +
"           and scc.own_corp_org_id = 3328\n" +
"         and scc.create_date >= trunc(sysdate)-7\n" +
"         and scc.create_date < trunc(sysdate)\n" +
"         and scc.cust_type = 4\n" +
"         and exists (select 1\n" +
"          from so1.Ins_Srvpkg sis\n" +
"         where sis.prod_service_id = 1002\n" +
"           and sis.state = 1\n" +
"           and sis.os_status is null\n" +
"           and sis.prod_inst_id = sip.prod_inst_id)\n" +
"         group by rad.jy_region_name,decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户')) tp1\n" +
" where tp.区域 <> '丁蜀镇'  and tp.区域 <>'其他'\n" +
"   and tp.区域 = tp1.区域(+)\n" +
" order by tp.区域,tp1.客户类型";

            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 2 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\周报\\" + "周报新增商业客户数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 周报月初欠停目前还欠停的带欠费金额及网格的客户明细
        public static bool zbaobiao16()
        {
            string sqlString =
"select scc.cust_code 客户证号,\n" +
"       scc.cust_name 客户姓名,\n" +
"       decode(scc.cust_type,\n" +
"              1,\n" +
"              '公众客户',\n" +
"              4,\n" +
"              '普通商业客户',\n" +
"              7,\n" +
"              '合同商业客户') 客户类型,\n" +
"       case\n" +
"         when exists\n" +
"          (select 1\n" +
"                 from so1.ins_prod            sip,\n" +
"                      so1.ins_prod_res        ipr,\n" +
"                      res.res_terminal        rrt,\n" +
"                      res.res_code_definition rcd\n" +
"                where sip.prod_inst_id = ipr.prod_inst_id\n" +
"                  and ipr.res_equ_no = rrt.serial_no\n" +
"                  and rrt.res_code = rcd.res_code\n" +
"                  and rcd.res_type = 2\n" +
"                  and (rcd.res_name like '%高清%' or rcd.res_name like '%4K%')\n" +
"                  and scc.cust_id = sip.cust_id) then\n" +
"          '有'\n" +
"         else\n" +
"          null\n" +
"       end 是否有高清或4K机顶盒,\n" +
"       con.cont_phone1 || ',' || con.cont_phone2 联系方式,\n" +
"       con.cont_mobile1 || ',' || con.cont_mobile2 移动电话,\n" +
"       wg.grid_name 所属网格,\n" +
"       rad.jy_region_name 区域,\n" +
"       iad.std_addr_name || iad.door_name 地址,\n" +
"       sum(zd.amount) / 100 欠费金额\n" +
"  from so1.cm_customer scc,\n" +
"       zg.acct zac,\n" +
"       so1.ins_address iad,\n" +
"       so1.cm_cust_contact_info con,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       wxjy.jy_customer_wg_rel wg,\n" +
"       (select *\n" +
"          from zg.acct_item_0\n" +
"        union all\n" +
"        select *\n" +
"          from zg.acct_item_1\n" +
"        union all\n" +
"        select *\n" +
"          from zg.acct_item_2\n" +
"        union all\n" +
"        select *\n" +
"          from zg.acct_item_3\n" +
"        union all\n" +
"        select *\n" +
"          from zg.acct_item_4\n" +
"        union all\n" +
"        select *\n" +
"          from zg.acct_item_5\n" +
"        union all\n" +
"        select *\n" +
"          from zg.acct_item_6\n" +
"        union all\n" +
"        select *\n" +
"          from zg.acct_item_7\n" +
"        union all\n" +
"        select *\n" +
"          from zg.acct_item_8\n" +
"        union all\n" +
"        select *\n" +
"          from zg.acct_item_9) zd\n" +
" where scc.cust_id = zac.cust_id\n" +
"   and zac.acct_id = zd.acct_id\n" +
"   and scc.cust_id = iad.cust_id\n" +
"   and scc.cust_id = con.cust_id\n" +
"   and scc.cust_id = rad.cust_id\n" +
"   and scc.cust_id = wg.cust_id(+)\n" +
"   and zd.amount > 0\n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and exists\n" +
" (select 1\n" +
"          from (select *\n" +
"                  from so1.ord_cust\n" +
"                union all\n" +
"                select *\n" +
"                  from so1.ord_cust_f_2020) a,\n" +
"               (select *\n" +
"                  from so1.ord_prod\n" +
"                union all\n" +
"                select *\n" +
"                  from so1.ord_prod_f_2020) b\n" +
"         where a.cust_order_id = b.cust_order_id\n" +
"           and a.business_id = '800001000077' --停机\n" +
"           and b.prod_spec_id = '800200000001' --数字电视\n" +
"           and a.create_date >= date '2020-" + DateTime.Now.ToString("MM") + "-1'\n" +
"           and a.create_date < date '2020-" + DateTime.Now.ToString("MM") + "-10'\n" +
"           and scc.cust_id = a.cust_id)\n" +
"   and exists (select 1\n" +
"          from so1.ins_prod sip, so1.ins_srvpkg sis\n" +
"         where sip.prod_inst_id = sis.prod_inst_id\n" +
"           and nvl(sip.main_prod_inst_id, 0) = 0\n" +
"           and sis.prod_service_id = 1002\n" +
"           and sis.os_status = '1'\n" +
"           and sip.cust_id = scc.cust_id)\n" +
" group by scc.cust_code,\n" +
"          scc.cust_name,\n" +
"          decode(scc.cust_type,\n" +
"                 1,\n" +
"                 '公众客户',\n" +
"                 4,\n" +
"                 '普通商业客户',\n" +
"                 7,\n" +
"                 '合同商业客户'),\n" +
"          case\n" +
"            when exists (select 1\n" +
"                    from so1.ins_prod            sip,\n" +
"                         so1.ins_prod_res        ipr,\n" +
"                         res.res_terminal        rrt,\n" +
"                         res.res_code_definition rcd\n" +
"                   where sip.prod_inst_id = ipr.prod_inst_id\n" +
"                     and ipr.res_equ_no = rrt.serial_no\n" +
"                     and rrt.res_code = rcd.res_code\n" +
"                     and rcd.res_type = 2\n" +
"                     and (rcd.res_name like '%高清%' or\n" +
"                         rcd.res_name like '%4K%')\n" +
"                     and scc.cust_id = sip.cust_id) then\n" +
"             '有'\n" +
"            else\n" +
"             null\n" +
"          end,\n" +
"          con.cont_phone1 || ',' || con.cont_phone2,\n" +
"          con.cont_mobile1 || ',' || con.cont_mobile2,\n" +
"          wg.grid_name,\n" +
"          rad.jy_region_name,\n" +
"          iad.std_addr_name || iad.door_name\n" +
" order by rad.jy_region_name";
            int[] columntxt = { 1,6 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\周报\\" + "月初欠停目前还欠停的带欠费金额及网格的客户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 无锡周报各项数据
        public static bool zbaobiao15()
        {
            string sqlString1 =
"select count(distinct  sip.prod_inst_id) 好视乐用户数,\n" +
"       count(distinct  scc.cust_id) 好视乐客户数\n" +
"  from so1.cm_customer            scc,\n" +
"       repnew.fact_ins_prod_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + "       sip,\n" +
"       repnew.fact_ins_srvpkg_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + "     sis,\n" +
"       product.up_product_item    upd\n" +
" where scc.cust_id = sip.cust_id\n" +
"   and sip.prod_inst_id = sis.prod_inst_id\n" +
"   and sis.srvpkg_id = upd.product_item_id\n" +
"    and upd.name like '好视乐%'\n" +
"   and scc.own_corp_org_id = 3328";
            DataTable dt1 = OracleHelper.ExecuteDataTable(sqlString1);//好视乐用户、客户数

            string sqlString2 =
"select count(distinct  sip.prod_inst_id)  看视界用户数,\n" +
"   count(distinct  scc.cust_id) 看视界客户数\n" +
"  from so1.cm_customer            scc,\n" +
"       repnew.fact_ins_prod_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + "       sip,\n" +
"       repnew.fact_ins_srvpkg_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + "     sis,\n" +
"       product.up_product_item    upd\n" +
"\n" +
" where scc.cust_id = sip.cust_id\n" +
"   and sip.prod_inst_id = sis.prod_inst_id\n" +
"   and sis.srvpkg_id = upd.product_item_id\n" +
"    and upd.name  in ('看视界组合包_56元/月_江阴',\n" +
"                     '看视界基础包_8元/月_江阴',\n" +
"                     '看视界69包_31元/月_江阴',\n" +
"                     '看视界89包_42元/月_江阴',\n" +
"                     '畅看D1_12元/月_江阴',\n" +
"                     '畅看D2_24元/月_江阴',\n" +
"                     '畅看B_0元/月_江阴')\n" +
"   and scc.own_corp_org_id = 3328";
            DataTable dt2 = OracleHelper.ExecuteDataTable(sqlString2);//看视界用户、客户数

            string sqlString3 =
" select count(distinct  sip.prod_inst_id) 广联用户数,\n" +
"   count(distinct  scc.cust_id) 广联客户数\n" +
" from so1.cm_customer            scc,\n" +
"      repnew.fact_ins_prod_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + "       sip,\n" +
"      repnew.fact_ins_srvpkg_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + "     sis,\n" +
"      product.up_product_item    upd\n" +
"where scc.cust_id = sip.cust_id\n" +
"  and sip.prod_inst_id = sis.prod_inst_id\n" +
"  and sis.offer_id= upd.product_item_id\n" +
" and upd.name like '广联合家欢%'\n" +
"  and scc.own_corp_org_id = 3328";
            DataTable dt3 = OracleHelper.ExecuteDataTable(sqlString3);//广联用户、客户数

            string sqlString4 =
" select count(distinct case\n" +
"              when (t1.state = '1' and t1.is_ins = '1' and t1.os_status is null or\n" +
"                   (t1.state = '99' or t1.is_ins = '0')) and\n" +
"                   pi.name not like '%测试%' and pi.name not like '%体验%' then\n" +
"               t1.prod_inst_id\n" +
"              else\n" +
"               null\n" +
"            end) 宽带缴费用户数\n" +
" from repnew.fact_Ins_prod_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + "   p,\n" +
"      repnew.fact_Ins_srvpkg_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " t1,\n" +
"      product.up_product_item         pi,\n" +
"      so1.cm_customer                 scc,\n" +
"      wxjy.jy_region_address_rel      rad\n" +
"where scc.cust_id = rad.cust_id\n" +
"  and p.cust_id = scc.cust_id\n" +
"  and p.prod_inst_id = t1.prod_inst_id\n" +
"  and t1.prod_service_id = 1004    \n" +
"  and p.corp_org_id = 3328\n" +
"  and t1.srvpkg_id = pi.product_item_id";
            DataTable dt4 = OracleHelper.ExecuteDataTable(sqlString4);//宽带缴费用户数

            string sqlString5 =
" select\n" +
"      count(distinct scc.cust_id) 留存缴费客户数\n" +
" from repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-01-01")).AddDays(-1).ToString("yyyyMMdd") + "       scc,\n" +
"      repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-01-01")).AddDays(-1).ToString("yyyyMMdd") + "   sip,\n" +
"      repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-01-01")).AddDays(-1).ToString("yyyyMMdd") + " sis\n" +
"where scc.cust_id = sip.cust_id\n" +
"  and sis.prod_inst_id = sip.prod_inst_id\n" +
"  and scc.own_corp_org_id = 3328\n" +
"  and sis.prod_service_id = 1002\n" +
"  and sis.state = 1\n" +
"  and sis.os_status is null\n" +
"  and exists (select 1\n" +
"         from repnew.fact_ins_srvpkg_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " sis1 \n" +
"        where sis1.prod_service_id = 1002\n" +
"          and sis1.state = 1\n" +
"          and sis1.os_status is null\n" +
"          and sis1.prod_inst_id = sip.prod_inst_id)";
            DataTable dt5 = OracleHelper.ExecuteDataTable(sqlString5);//留存缴费客户数

            string sqlString6 =
"select\n" +
"       count(distinct sip.prod_inst_id) 江阴互动基本缴费用户数\n" +
"  from repnew.fact_ins_prod_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " sip, repnew.fact_cust_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " scc   --------\n" +
" where sip.cust_id = scc.cust_id\n" +
"   and sip.is_paied = 1\n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and exists\n" +
" (select 1\n" +
"          from repnew.fact_ins_srvpkg_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " sis1     -------\n" +
"         where exists (select 1\n" +
"                  from product.up_product_item upd1\n" +
"                 where upd1.name = '江阴本地回看'\n" +
"                   and upd1.product_item_id = sis1.srvpkg_id)\n" +
"           and sip.prod_inst_id = sis1.prod_inst_id)\n" +
"   and exists\n" +
" (select 1\n" +
"          from repnew.fact_ins_srvpkg_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " sis2     -------\n" +
"         where exists (select 1\n" +
"                  from product.up_product_item upd2\n" +
"                 where upd2.name = '栏目回看'\n" +
"                   and upd2.product_item_id = sis2.srvpkg_id)\n" +
"           and sip.prod_inst_id = sis2.prod_inst_id)\n" +
"   and exists\n" +
" (select 1\n" +
"          from repnew.fact_ins_srvpkg_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " sis3      -------\n" +
"         where exists (select 1\n" +
"                  from product.up_product_item upd3\n" +
"                 where upd3.name = '频道回看'\n" +
"                   and upd3.product_item_id = sis3.srvpkg_id)\n" +
"           and sip.prod_inst_id = sis3.prod_inst_id)\n" +
"   and exists\n" +
" (select 1\n" +
"          from repnew.fact_ins_srvpkg_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " sis4     ------\n" +
"         where exists (select 1\n" +
"                  from product.up_product_item upd4\n" +
"                 where upd4.name = '时移'\n" +
"                   and upd4.product_item_id = sis4.srvpkg_id)\n" +
"           and sip.prod_inst_id = sis4.prod_inst_id)";
            DataTable dt6 = OracleHelper.ExecuteDataTable(sqlString6);//江阴互动基本缴费用户数

            string sqlString7 =
"select\n" +
"      count(distinct scc.cust_code) 周新增客户数\n" +
" from repnew.fact_cust_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " scc\n" +
"where scc.create_date >= date '" + DateTime.Now.AddDays(-7).ToString("yyyy-MM-dd") + "'\n" +
"  and scc.create_date < date '" + DateTime.Now.ToString("yyyy-MM-dd") + "'\n" +
"  and scc.cust_type = 1\n" +
"  and scc.own_corp_org_id = 3328\n" +
"  and exists (select 1\n" +
"         from repnew.fact_ins_prod_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + "   sip,\n" +
"              repnew.fact_ins_srvpkg_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " sis\n" +
"        where sip.prod_inst_id = sis.prod_inst_id\n" +
"          and sis.prod_service_id = 1002\n" +
"          and sis.state = 1\n" +
"          and sis.os_status is null)";
            DataTable dt7 = OracleHelper.ExecuteDataTable(sqlString7);//周新增客户数

            string sqlString8 =
"select count(distinct scc.cust_code) 年新增客户数\n" +
"  from repnew.fact_cust_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " scc\n" +
" where scc.create_date >= date '" + DateTime.Now.ToString("yyyy-01-01") + "'\n" +
"   and scc.create_date < date '" + DateTime.Now.ToString("yyyy-MM-dd") + "'\n" +
"   and scc.cust_type = 1\n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and exists (select 1\n" +
"          from repnew.fact_ins_prod_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + "   sip,\n" +
"               repnew.fact_ins_srvpkg_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " sis\n" +
"         where sip.prod_inst_id = sis.prod_inst_id\n" +
"           and sis.prod_service_id = 1002\n" +
"           and sis.state = 1\n" +
"           and sis.os_status is null\n" +
"           and sip.cust_id = scc.cust_id)\n" +
"   and not exists (select 1\n" +
"          from repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-01-01")).AddDays(-1).ToString("yyyyMMdd") + " scc1\n" +
"         where scc1.cust_id = scc.cust_id)";
            DataTable dt8 = OracleHelper.ExecuteDataTable(sqlString8);//年新增客户数

            string sqlString9 =
"with t_all as\n" +
"    (select t.cust_id,t.corp_org_id,\n" +
"       MAX(CASE WHEN t.is_dtv = 1 THEN NVL(t.boder_flag,1)\n" +
"              WHEN t.is_atv = 1 THEN -1\n" +
"              END\n" +
"       ) tclx,\n" +
"       MIN(case when t.user_status = 1 then 2 when t.user_status = 2 then 1 else t.user_status end\n" +
"       ) user_status\n" +
"    from repnew.fact_ins_Prod_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " t\n" +
"    left join repnew.fact_yunweibujiqi_wx t3 on t.cust_id = t3.cust_id\n" +
"    WHERE 1=1 AND (t.is_dtv = 1 OR t.is_atv = 1)\n" +
"        and t3.cust_id is null\n" +
"    GROUP BY t.cust_id,t.corp_org_id\n" +
"    ),\n" +
" t_cancel as\n" +
"    (SELECT  t.cust_id,t.corp_org_id,\n" +
"    MAX(CASE WHEN t.is_dtv = 1 THEN NVL(t.boder_flag,1)\n" +
"              WHEN t.is_atv = 1 THEN -1\n" +
"              END\n" +
"       ) tclx,\n" +
"       MIN(case when t.user_status = 1 then 2 when t.user_status = 2 then 1 else t.user_status end\n" +
"       ) user_status\n" +
"    from repnew.fact_ins_prod_cancel t\n" +
"    where (t.is_dtv=1 OR t.is_atv=1)\n" +
"        AND NOT EXISTS (select 1 from repnew.fact_ins_Prod_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " fip WHERE t.cust_id = fip.cust_id)\n" +
"    GROUP BY t.cust_id,t.corp_org_id\n" +
"    )\n" +
"\n" +
"    select\n" +
"       sum(tt.ktkhs) 开通客户数,\n" +
"       sum(tt.yktkhs) 预开通客户数\n" +
"       from\n" +
"(\n" +
"  select\n" +
"             t1.corp_org_id,\n" +
"             count(distinct case WHEN t1.user_status = 1 then t1.cust_id else null end) ktkhs, --开通客户数\n" +
"             count(distinct case when t1.user_status = 2 then t1.cust_id else null end) yktkhs --预开通客户数\n" +
"        from t_all t1\n" +
"        LEFT JOIN repnew.fact_cust_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " cust ON t1.cust_id = cust.cust_id\n" +
"        join wxjy.jy_region_address_rel dz on t1.cust_id=dz.cust_id\n" +
"       group by t1.corp_org_id,dz.jy_region_name,\n" +
"                cust.cust_type\n" +
"\n" +
"\n" +
"    UNION ALL\n" +
"    select t1.corp_org_id,\n" +
"             0 ktkhs, --开通客户数\n" +
"             0 yktkhs --预开通客户数\n" +
"        from t_cancel t1\n" +
"        LEFT JOIN repnew.fact_cust_" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + " cust ON t1.cust_id = cust.cust_id\n" +
"                join  wxjy.jy_region_address_rel dz on t1.cust_id=dz.cust_id\n" +
"       group by t1.corp_org_id,dz.jy_region_name,t1.tclx, cust.cust_type\n" +
"               ) tt\n" +
"where tt.corp_org_id = 3328";
            DataTable dt9 = OracleHelper.ExecuteDataTable(sqlString9);//数字电视缴费客户数
            ExcelHelper.DataTableTowuxiExcel("\\周报\\" + "无锡周报各项数据" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt1, dt2, dt3, dt4, dt5, dt6, dt7, dt8, dt9, sqlString1+ sqlString2 + sqlString3 + sqlString4 + sqlString5 + sqlString6 + sqlString7 + sqlString8 + sqlString9);
            return true;
        }
        #endregion


        


        #region 月报720套餐相关发票的开票明细数据
        public static bool ybaobiao1()
        {
            string sqlString =

"select a.invoice_list_no 发票代码,\n" +
"       a.invoice_num 发票号码,\n" +
"       a.invoice_amount / 100 发票金额,\n" +
"       a.image 科目,\n" +
"       decode(a.invoice_state, 1, '正常', 3, '冲红', 4, '被冲红', '') 发票类型,\n" +
"       a.print_date 开票时间,\n" +
"       a.remark 备注,\n" +
"       c.cust_code 客户证号,\n" +
"       c.cust_name 客户姓名,\n" +
"       b.staff_name 开票人,\n" +
"       e.organize_name 营业厅\n" +
"  from so1.ord_invoice_2019 a,\n" +
"       so1.cm_customer      c,\n" +
"       sec.sec_staff        b,\n" +
"       sec.sec_operator     d,\n" +
"       sec.sec_organize     e\n" +
" where a.cust_id = c.cust_id\n" +
"   and c.own_corp_org_id = 3328\n" +
"   and a.op_id = d.operator_id\n" +
"   and d.staff_id = b.staff_id\n" +
"   and a.org_id = e.organize_id\n" +
"   and a.remark like '%720%'\n" +
"   and a.print_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"   and a.print_date < date '"+ DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
" order by a.invoice_num";

            int[] columntxt = { 1, 2, 8 }; //哪些列是文本格式
            int[] columndate = { 6 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "720套餐相关发票的开票明细数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报360元分2和5年及720元分3和5年返还的客户返充明细数据
        public static bool ybaobiao2()
        {
            string sqlString =

"select distinct scc.cust_code 客户证号,\n" +
"                scc.cust_name 客户姓名,\n" +
"                a.balance_log_id 流水ID,\n" +
"                a.amount / 100 返充金额,\n" +
"                a.payment_date 返充时间,\n" +
"                rad.jy_region_name 区域,\n" +
"                iad.std_addr_name || iad.door_name 客户地址\n" +
"  from zg.acct_balance_log_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " a, ---账本流水记录表：统计哪个月的数据就需改成哪个月的表\n" +
"       so1.cm_customer            scc,\n" +
"       zg.acct                    zac,\n" +
"       zg.acct_balance            zab,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       so1.ins_address            iad\n" +
" where a.acct_balance_id = zab.acct_balance_id\n" +
"   and zab.acct_id = zac.acct_id\n" +
"   and zac.cust_id = scc.cust_id\n" +
"   and scc.cust_id = iad.cust_id\n" +
"   and scc.cust_id = rad.cust_id\n" +
"   and zab.balance_type_id = 1\n" +
"   and (a.amount in (1200, -1200) or a.amount in (2000, -2000) or\n" +
"       a.amount in (600, -600) or a.amount in (1500, -1500) or a.amount in (10000, -10000))\n" +
"   and a.operation_type = 141000\n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and scc.cust_prop <> 4\n" +
" order by scc.cust_code";
            int[] columntxt = { 1 }; //哪些列是文本格式
            int[] columndate = { 5 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "360元分2和5年及720元分3和5年返还及押金返点的客户返充明细数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报订购720元分3、5年返还套餐客户明细
        public static bool ybaobiao3()
        {
            string sqlString =

"select scc.cust_code     客户证号,\n" +
"       scc.cust_name     客户姓名,\n" +
"       zpd.remark        续费返充活动,\n" +
"       sof.create_date   订购时间,\n" +
"       sta.staff_name    操作员,\n" +
"       sor.organize_name 营业厅\n" +
"  from so1.cm_customer          scc,\n" +
"       so1.ins_prod             sip,\n" +
"       so1.ins_off_ins_prod_rel rel,\n" +
"       so1.ins_offer            sof,\n" +
"       zg.acct                  zac,\n" +
"       zg.acc_pa_allot_info     zpa,\n" +
"       zg.acc_pa_allot_rule_def zpd,\n" +
"       sec.sec_staff            sta,\n" +
"       sec.sec_operator         sop,\n" +
"       sec.sec_organize         sor\n" +
" where scc.cust_id = sip.cust_id\n" +
"   and sip.prod_inst_id = rel.prod_inst_id\n" +
"   and rel.offer_inst_id = sof.offer_inst_id\n" +
"   and zpd.allot_rule_id = zpa.promo_id\n" +
"   and sof.op_id = sop.operator_id\n" +
"   and sop.staff_id = sta.staff_id\n" +
"   and sof.org_id = sor.organize_id\n" +
"   and scc.cust_id = zac.cust_id\n" +
"   and zac.acct_id = zpa.acct_id\n" +
"   and sof.offer_id in (800056,800057)\n" +
"   and zpd.allot_rule_id in(1000056,1000057)\n" +
"   and sof.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"   and sof.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"   and zpa.so_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"   and zpa.so_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
" order by scc.cust_code";

            int[] columntxt = { 1 }; //哪些列是文本格式
            int[] columndate = { 4 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "订购720元分3、5年返还套餐客户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报360套餐相关发票的开票明细数据
        public static bool ybaobiao4()
        {
            string sqlString =

"select a.invoice_list_no 发票代码,\n" +
"       a.invoice_num 发票号码,\n" +
"       a.invoice_amount / 100 发票金额,\n" +
"       a.image 科目,\n" +
"       decode(a.invoice_state, 1, '正常', 3, '冲红', 4, '被冲红', '') 发票类型,\n" +
"       a.print_date 开票时间,\n" +
"       a.remark 备注,\n" +
"       c.cust_code 客户证号,\n" +
"       c.cust_name 客户姓名,\n" +
"       b.staff_name 开票人,\n" +
"       e.organize_name 营业厅\n" +
"  from so1.ord_invoice_2020 a,\n" +
"       so1.cm_customer      c,\n" +
"       sec.sec_staff        b,\n" +
"       sec.sec_operator     d,\n" +
"       sec.sec_organize     e\n" +
" where a.cust_id = c.cust_id\n" +
"   and c.own_corp_org_id = 3328\n" +
"   and a.op_id = d.operator_id\n" +
"   and d.staff_id = b.staff_id\n" +
"   and a.org_id = e.organize_id\n" +
"   and (a.remark like '%360%' or a.remark like '%4K%')\n" +
"   and a.print_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"   and a.print_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
" order by a.invoice_num";

            int[] columntxt = { 1, 2, 8 }; //哪些列是文本格式
            int[] columndate = { 6 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "360套餐相关发票的开票明细数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报订购三务公开高清互动升级充值360元分5和2年返还套餐客户明细
        public static bool ybaobiao5()
        {
            string sqlString =
"select distinct  scc.cust_code     客户证号,\n" +
"       scc.cust_name     客户姓名,\n" +
"       zpd.promo_name    续费返充活动,\n" +
"       sof.create_date   订购时间,\n" +
"       sta.staff_name    操作员,\n" +
"       sof.offer_inst_id,\n" +
"       sor.organize_name 营业厅\n" +
"  from so1.cm_customer          scc,\n" +
"       so1.ins_prod             sip,\n" +
"       so1.ins_off_ins_prod_rel rel,\n" +
"       so1.ins_offer            sof,\n" +
"       sec.sec_staff            sta,\n" +
"       sec.sec_operator         sop,\n" +
"       sec.sec_organize         sor,\n" +
"       zg.acct                  zac,\n" +
"       zg.ACC_PA_ALLOT_INFO     zpa,\n" +
"       zg.acc_pa_promo_def      zpd\n" +
" where scc.cust_id = sip.cust_id\n" +
"   and sip.prod_inst_id = rel.prod_inst_id\n" +
"   and rel.offer_inst_id = sof.offer_inst_id\n" +
"   and sof.op_id = sop.operator_id\n" +
"   and sop.staff_id = sta.staff_id\n" +
"   and sof.org_id = sor.organize_id\n" +
"   and zac.cust_id = scc.cust_id\n" +
"   and zpa.acct_id = zac.acct_id\n" +
"   and zpa.serv_id = sip.prod_inst_id\n" +
"   and zpd.promo_id = zpa.promo_id\n" +
"   and sof.offer_id = zpd.outer_promo_id\n" +
"   and zpd.promo_id in (1000055, 1000060, 1000069, 1000070)\n" +
"   and sof.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"   and sof.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"   and scc.cust_prop <> 4\n" +
"   and scc.own_corp_org_id = 3328";
            int[] columntxt = { 1,6 }; //哪些列是文本格式
            int[] columndate = { 4 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "订购三务公开高清互动升级充值360元分5和2年返还及押金返点套餐客户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报云亭纯标清缴费客户明细
        public static bool ybaobiao7()
        {
            string sqlString =
"select distinct scc.cust_code 客户证号,\n" +
"                scc.cust_name 客户姓名,\n" +
"                decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户') 客户类型,\n" +
"                scc.mobile2 || ',' || scc.mobile2 联系方式1,\n" +
"                scc.phone1 || ',' || scc.phone2 联系方式2,\n" +
"                trim(case\n" +
"                       when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                            scc.std_addr_name not like '%测试%' then\n" +
"                        substr(scc.std_addr_name, 13, 3)\n" +
"                       when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                            scc.std_addr_name not like '%测试%' then\n" +
"                        substr(scc.std_addr_name, 6, 3)\n" +
"                       when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                        '澄江'\n" +
"                       when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                        '华士'\n" +
"                       when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                        '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                       when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                        '祝塘'\n" +
"                       when (scc.std_addr_name like '%测试%' or\n" +
"                            scc.std_addr_name is null) then\n" +
"                        '其他'\n" +
"                       else\n" +
"                        substr(scc.std_addr_name, 1, 3)\n" +
"                     end) 区域,\n" +
"                scc.std_addr_name 地址\n" +
"  from repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " scc\n" +
" where scc.own_corp_org_id = 3328\n" +
"   and trim(case\n" +
"              when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 13, 3)\n" +
"              when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 6, 3)\n" +
"              when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"               '澄江'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"               '华士'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"               '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"               '祝塘'\n" +
"              when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"               '其他'\n" +
"              else\n" +
"               substr(scc.std_addr_name, 1, 3)\n" +
"            end) = '云亭'\n" +
"   and exists (select 1\n" +
"          from repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sis,\n" +
"               repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "   sip\n" +
"         where sip.prod_inst_id = sis.prod_inst_id\n" +
"           and sis.prod_service_id = 1002\n" +
"           and sis.state = 1\n" +
"           and sis.os_status is null\n" +
"           and sip.cust_id = scc.cust_id)\n" +
"   and not exists\n" +
" (select 1\n" +
"          from repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sip1,\n" +
"               res.res_terminal              rrt,\n" +
"               res.res_code_definition       rcd\n" +
"         where sip1.stb_id = rrt.serial_no\n" +
"           and rrt.res_code = rcd.res_code\n" +
"           and (rcd.res_name like '%高清%' or rcd.res_name like '%4K%')\n" +
"           and sip1.cust_id = scc.cust_id) --不存在高清4K用户";
            int[] columntxt = { 1 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "云亭纯标清缴费客户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报云亭缴费客户明细
        public static bool ybaobiao8()
        {
            string sqlString =
"select distinct scc.cust_code 客户证号,\n" +
"                scc.cust_name 客户姓名,\n" +
"                decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户') 客户类型,\n" +
"                scc.mobile2 || ',' || scc.mobile2 联系方式1,\n" +
"                scc.phone1 || ',' || scc.phone2 联系方式2,\n" +
"                trim(case\n" +
"                       when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                            scc.std_addr_name not like '%测试%' then\n" +
"                        substr(scc.std_addr_name, 13, 3)\n" +
"                       when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                            scc.std_addr_name not like '%测试%' then\n" +
"                        substr(scc.std_addr_name, 6, 3)\n" +
"                       when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                        '澄江'\n" +
"                       when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                        '华士'\n" +
"                       when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                        '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                       when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                        '祝塘'\n" +
"                       when (scc.std_addr_name like '%测试%' or\n" +
"                            scc.std_addr_name is null) then\n" +
"                        '其他'\n" +
"                       else\n" +
"                        substr(scc.std_addr_name, 1, 3)\n" +
"                     end) 区域,\n" +
"                scc.std_addr_name 地址\n" +
"  from repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " scc\n" +
" where scc.own_corp_org_id = 3328\n" +
"   and trim(case\n" +
"              when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 13, 3)\n" +
"              when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 6, 3)\n" +
"              when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"               '澄江'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"               '华士'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"               '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"               '祝塘'\n" +
"              when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"               '其他'\n" +
"              else\n" +
"               substr(scc.std_addr_name, 1, 3)\n" +
"            end) = '云亭'\n" +
"   and exists (select 1\n" +
"          from repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sis,\n" +
"               repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "   sip\n" +
"         where sip.prod_inst_id = sis.prod_inst_id\n" +
"           and sis.prod_service_id = 1002\n" +
"           and sis.state = 1\n" +
"           and sis.os_status is null\n" +
"           and sip.cust_id = scc.cust_id)";
            int[] columntxt = { 1 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "云亭缴费客户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报新增宽带用户数据统计
        public static bool ybaobiao9()
        {
            string sqlString =

"select rad.jy_region_name 区域, count(distinct op.prod_inst_id) 新增用户数\n" +
"  from (select *\n" +
"          from so1.ord_cust\n" +
"        union all\n" +
"        select *\n" +
"          from so1.ord_cust_f_2020) oc,\n" +
"       (select *\n" +
"          from so1.ord_prod\n" +
"        union all\n" +
"        select *\n" +
"          from so1.ord_prod_f_2020) op,\n" +
"       (select *\n" +
"          from so1.ord_srvpkg\n" +
"        union all\n" +
"        select *\n" +
"          from so1.ord_srvpkg_f_2020) os,\n" +
"       so1.cm_customer scc,\n" +
"       so1.ins_address iad,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where scc.cust_id = op.cust_id\n" +
"   and scc.cust_id = iad.cust_id(+)\n" +
"   and scc.cust_id = rad.cust_id\n" +
"   and oc.CUST_ORDER_ID = op.CUST_ORDER_ID\n" +
"   and op.prod_order_id = os.prod_order_id\n" +
"   and oc.BUSINESS_ID in (800001000001, 800001000002) --普通新装,批量新装\n" +
"   and op.prod_spec_id = 800200000003 --宽带:800200000003   数字:800200000001\n" +
"   and oc.order_state <> 10 --排除新装撤单情况\n" +
"   and oc.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"   and oc.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"   and oc.own_corp_org_id = 3328\n" +
" group by rad.jy_region_name";

            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "新增宽带用户数据统计--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报新增高清机顶盒用户明细
        public static bool ybaobiao10()
        {
            string sqlString =

"select distinct scc.cust_code 客户证号,\n" +
"                scc.cust_name 客户姓名,\n" +
"                decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户') 客户类型,\n" +
"                con.cont_phone1 || ',' || con.cont_phone2 联系方式1,\n" +
"                con.cont_mobile1 || ',' || con.cont_mobile2 联系方式2,\n" +
"                ipr.res_equ_no 机顶盒号码,\n" +
"                rad.jy_region_name 区域,\n" +
"                iad.std_addr_name || ',' || iad.door_name 地址\n" +
"  from so1.ins_prod_res           ipr,\n" +
"       so1.ins_prod               sip,\n" +
"       so1.cm_customer            scc,\n" +
"       so1.cm_cust_contact_info   con,\n" +
"       so1.ins_address            iad,\n" +
"       res.res_code_definition    rcd,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where ipr.res_equ_no = sip.bill_id\n" +
"   and sip.cust_id = scc.cust_id\n" +
"   and scc.cust_id = con.cust_id(+)\n" +
"   and scc.cust_id = iad.cust_id(+)\n" +
"   and scc.cust_id = rad.cust_id\n" +
"   and ipr.res_type = rcd.res_sub_type\n" +
"   and (rcd.res_name like '%高清%' or rcd.res_name like '%4K%')\n" +
"   and sip.sub_bill_id is not null\n" +
"   and ipr.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"   and ipr.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"   and scc.own_corp_org_id = 3328\n" +
" order by scc.cust_code";
            int[] columntxt = { 1,6 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "新增高清机顶盒用户明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报数字有效互动客户、用户数量
        public static bool ybaobiao11()
        {
            string sqlString =
"select rad.jy_region_name 区域,\n" +
"       count(distinct sip.prod_inst_id) 有效互动用户数,\n" +
"       count(distinct cust.cust_id) 有效互动客户数\n" +
"  from repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "       cust,\n" +
"       repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "   sip,\n" +
"       repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sis,\n" +
"       wxjy.jy_region_address_rel      rad\n" +
" where cust.cust_id = rad.cust_id\n" +
"   and cust.cust_id = sip.cust_id\n" +
"   and sip.prod_inst_id = sis.prod_inst_id\n" +
"   and sip.is_dtv = 1\n" +
"   and sip.is_valid1 = 1\n" +
"   and cust.own_corp_org_id = 3328\n" +
"   and exists (select 1\n" +
"          from wxjy.rep_jy_itv_product t2\n" +
"         where t2.product_item_id = sis.srvpkg_id)\n" +
" group by rad.jy_region_name\n" +
" order by rad.jy_region_name";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "数字有效互动客户、用户数量--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报数字有效客户数
        public static bool ybaobiao12()
        {
            string sqlString =
"select rad.jy_region_name 区域,\n" +
"       count(distinct cust.cust_id) 数字电视有效客户数\n" +
"  from repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "     cust,\n" +
"       repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sip,\n" +
"       wxjy.jy_region_address_rel    rad\n" +
" where cust.cust_id = rad.cust_id\n" +
"   and cust.cust_id = sip.cust_id\n" +
"   and sip.is_dtv = 1\n" +
"   and sip.is_valid1 = 1\n" +
"   and cust.own_corp_org_id = 3328\n" +
" group by rad.jy_region_name\n" +
" order by rad.jy_region_name";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "数字有效客户数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报数字缴费等用户数
        public static bool ybaobiao13()
        {
            string sqlString =
"with t_all as\n" +
"    (select t.cust_id,t.corp_org_id,t.prod_inst_id,\n" +
"       MAX(CASE WHEN t.is_dtv = 1 THEN NVL(t.boder_flag,1)\n" +
"              WHEN t.is_atv = 1 THEN -1\n" +
"              END\n" +
"       ) tclx,\n" +
"       MIN(case when t.user_status = 1 then 2 when t.user_status = 2 then 1 else t.user_status end\n" +
"       ) user_status\n" +
"    from repnew.fact_ins_Prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " t ----1\n" +
"    left join repnew.fact_yunweibujiqi_wx t3 on t.cust_id = t3.cust_id\n" +
"    WHERE 1=1 AND (t.is_dtv = 1 OR t.is_atv = 1)\n" +
"        and t3.cust_id is null\n" +
"    GROUP BY t.cust_id,t.corp_org_id,t.prod_inst_id\n" +
"    ),\n" +
" t_cancel as\n" +
"    (SELECT  t.cust_id,t.corp_org_id,t.prod_inst_id,\n" +
"    MAX(CASE WHEN t.is_dtv = 1 THEN NVL(t.boder_flag,1)\n" +
"              WHEN t.is_atv = 1 THEN -1\n" +
"              END\n" +
"       ) tclx,\n" +
"       MIN(case when t.user_status = 1 then 2 when t.user_status = 2 then 1 else t.user_status end\n" +
"       ) user_status\n" +
"    from repnew.fact_ins_prod_cancel t\n" +
"    where (t.is_dtv=1 OR t.is_atv=1)\n" +
"        AND NOT EXISTS (select 1 from repnew.fact_ins_Prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " fip WHERE t.cust_id = fip.cust_id)   ----2\n" +
"    GROUP BY t.cust_id,t.corp_org_id,t.prod_inst_id\n" +
"    )\n" +
"    select\n" +
"     tt.corp_org_id,\n" +
"     tt.区域,tt.非行政区域,\n" +
"     tt.tclx,\n" +
"     tt.cust_type,\n" +
"       sum(tt.ktkhs) 开通用户数,\n" +
"       sum(tt.yktkhs) 预开通用户数,\n" +
"       sum(tt.ztkhs) 暂停用户数,\n" +
"       sum(tt.qftjkhs) 欠费停机用户数,\n" +
"       sum(tt.qfjxkhs) 欠费剪线用户数,\n" +
"       sum(tt.yxhkhs) 预销户用户数,\n" +
"       sum(tt.xhkhs) 销户用户数\n" +
"       from\n" +
"(\n" +
"  select\n" +
"             t1.corp_org_id,\n" +
"             dz.jy_region_name 区域,dz.jy_non_region_name 非行政区域,\n" +
"             decode(t1.tclx, 1, '境内', 2, '境外', -1,'模拟','境内') tclx,\n" +
"             DECODE(cust.cust_type,1,'公众客户',4,'普通商业客户',7,'合同商业客户','公众客户') cust_type,\n" +
"             count(distinct case WHEN t1.user_status = 1 then t1.prod_inst_id else null end) ktkhs, --开通用户数\n" +
"             count(distinct case when t1.user_status = 2 then t1.prod_inst_id else null end) yktkhs, --预开通用户数\n" +
"             count(distinct case when t1.user_status = 6 then t1.prod_inst_id else null end) ztkhs, --暂停用户数\n" +
"             count(distinct case when t1.user_status in( 3,9) then t1.prod_inst_id else null end) qftjkhs, --欠费停机用户数\n" +
"             count(distinct case when t1.user_status = 4 then t1.prod_inst_id else null end) qfjxkhs, --欠费剪线用户数\n" +
"             COUNT(DISTINCT CASE WHEN t1.user_status=7 THEN t1.prod_inst_id ELSE NULL END) yxhkhs, --预销户用户数\n" +
"             COUNT(DISTINCT CASE WHEN t1.user_status=8 THEN t1.prod_inst_id ELSE NULL END)xhkhs --销户用户数\n" +
"        from t_all t1\n" +
"        LEFT JOIN repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " cust ON t1.cust_id = cust.cust_id  ----3\n" +
"        join  wxjy.jy_region_address_rel dz on t1.cust_id=dz.cust_id\n" +
"       group by t1.corp_org_id,dz.jy_region_name,dz.jy_non_region_name,\n" +
"                 t1.tclx,\n" +
"                cust.cust_type\n" +
"    UNION ALL\n" +
"    select t1.corp_org_id,\n" +
" dz.jy_region_name 区域,dz.jy_non_region_name 非行政区域,\n" +
"             decode(t1.tclx, 1, '境内', 2, '境外', -1,'模拟','境内') tclx,\n" +
"             DECODE(cust.cust_type,1,'公众客户',4,'普通商业客户',7,'合同商业客户','公众客户') cust_type,\n" +
"             0 ktkhs, --开通客户数\n" +
"             0 yktkhs, --预开通客户数\n" +
"             0 ztkhs, --暂停客户数\n" +
"             0 qftjkhs, --欠费停机客户数\n" +
"             0 qfjxkhs, --欠费剪线客户数\n" +
"             COUNT(DISTINCT CASE WHEN t1.user_status=7 THEN t1.prod_inst_id ELSE NULL END) yxhkhs, --预销户客户数\n" +
"             COUNT(DISTINCT CASE WHEN t1.user_status=8 THEN t1.prod_inst_id ELSE NULL END)xhkhs --销户客户数\n" +
"        from t_cancel t1\n" +
"        LEFT JOIN repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " cust ON t1.cust_id = cust.cust_id  ----4\n" +
"                join  wxjy.jy_region_address_rel dz on t1.cust_id=dz.cust_id\n" +
"       group by t1.corp_org_id,dz.jy_region_name, dz.jy_non_region_name, t1.tclx, cust.cust_type\n" +
"               ) tt\n" +
"where tt.corp_org_id = 3328\n" +
" GROUP BY tt.corp_org_id,tt.区域,tt.非行政区域,tt.tclx,tt.cust_type";

            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "数字缴费等用户数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报数字缴费等客户数
        public static bool ybaobiao14()
        {
            string sqlString =
"with t_all as\n" +
"    (select t.cust_id,t.corp_org_id,\n" +
"       MAX(CASE WHEN t.is_dtv = 1 THEN NVL(t.boder_flag,1)\n" +
"              WHEN t.is_atv = 1 THEN -1\n" +
"              END\n" +
"       ) tclx,\n" +
"       MIN(case when t.user_status = 1 then 2 when t.user_status = 2 then 1 else t.user_status end\n" +
"       ) user_status\n" +
"    from repnew.fact_ins_Prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " t               --------1\n" +
"    left join repnew.fact_yunweibujiqi_wx t3 on t.cust_id = t3.cust_id\n" +
"    WHERE 1=1 AND (t.is_dtv = 1 OR t.is_atv = 1)\n" +
"        and t3.cust_id is null\n" +
"    GROUP BY t.cust_id,t.corp_org_id\n" +
"    ),\n" +
" t_cancel as\n" +
"    (SELECT  t.cust_id,t.corp_org_id,\n" +
"    MAX(CASE WHEN t.is_dtv = 1 THEN NVL(t.boder_flag,1)\n" +
"              WHEN t.is_atv = 1 THEN -1\n" +
"              END\n" +
"       ) tclx,\n" +
"       MIN(case when t.user_status = 1 then 2 when t.user_status = 2 then 1 else t.user_status end\n" +
"       ) user_status\n" +
"    from repnew.fact_ins_prod_cancel t\n" +
"    where (t.is_dtv=1 OR t.is_atv=1)\n" +
"        AND NOT EXISTS (select 1 from repnew.fact_ins_Prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " fip WHERE t.cust_id = fip.cust_id)   --------2\n" +
"    GROUP BY t.cust_id,t.corp_org_id\n" +
"    )\n" +
"    select\n" +
"     tt.corp_org_id,\n" +
"     tt.区域,tt.非行政区域,\n" +
"     tt.tclx,\n" +
"     tt.cust_type,\n" +
"       sum(tt.ktkhs) 开通客户数,\n" +
"       sum(tt.yktkhs) 预开通客户数,\n" +
"       sum(tt.ztkhs) 暂停客户数,\n" +
"       sum(tt.qftjkhs) 欠费停机客户数,\n" +
"       sum(tt.qfjxkhs) 欠费剪线客户数,\n" +
"       sum(tt.yxhkhs) 预销户客户数,\n" +
"       sum(tt.xhkhs) 销户客户数\n" +
"       from\n" +
"(\n" +
"  select\n" +
"             t1.corp_org_id,\n" +
"             dz.jy_region_name 区域,dz.jy_non_region_name 非行政区域,\n" +
"             decode(t1.tclx, 1, '境内', 2, '境外', -1,'模拟','境内') tclx,\n" +
"             DECODE(cust.cust_type,1,'公众客户',4,'普通商业客户',7,'合同商业客户','公众客户') cust_type,\n" +
"             count(distinct case WHEN t1.user_status = 1 then t1.cust_id else null end) ktkhs, --开通客户数\n" +
"             count(distinct case when t1.user_status = 2 then t1.cust_id else null end) yktkhs, --预开通客户数\n" +
"             count(distinct case when t1.user_status = 6 then t1.cust_id else null end) ztkhs, --暂停客户数\n" +
"             count(distinct case when t1.user_status in( 3,9) then t1.cust_id else null end) qftjkhs, --欠费停机客户数\n" +
"             count(distinct case when t1.user_status = 4 then t1.cust_id else null end) qfjxkhs, --欠费剪线客户数\n" +
"             COUNT(DISTINCT CASE WHEN t1.user_status=7 THEN t1.cust_id ELSE NULL END) yxhkhs, --预销户客户数\n" +
"             COUNT(DISTINCT CASE WHEN t1.user_status=8 THEN t1.cust_id ELSE NULL END)xhkhs --销户客户数\n" +
"        from t_all t1\n" +
"        LEFT JOIN repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " cust ON t1.cust_id = cust.cust_id    --------3\n" +
"        join wxjy.jy_region_address_rel dz on t1.cust_id=dz.cust_id\n" +
"       group by t1.corp_org_id,dz.jy_region_name,dz.jy_non_region_name,\n" +
"                 t1.tclx,\n" +
"                cust.cust_type\n" +
"    UNION ALL\n" +
"    select t1.corp_org_id,\n" +
" dz.jy_region_name 区域,dz.jy_non_region_name 非行政区域,\n" +
"             decode(t1.tclx, 1, '境内', 2, '境外', -1,'模拟','境内') tclx,\n" +
"             DECODE(cust.cust_type,1,'公众客户',4,'普通商业客户',7,'合同商业客户','公众客户') cust_type,\n" +
"             0 ktkhs, --开通客户数\n" +
"             0 yktkhs, --预开通客户数\n" +
"             0 ztkhs, --暂停客户数\n" +
"             0 qftjkhs, --欠费停机客户数\n" +
"             0 qfjxkhs, --欠费剪线客户数\n" +
"             COUNT(DISTINCT CASE WHEN t1.user_status=7 THEN t1.cust_id ELSE NULL END) yxhkhs, --预销户客户数\n" +
"             COUNT(DISTINCT CASE WHEN t1.user_status=8 THEN t1.cust_id ELSE NULL END)xhkhs --销户客户数\n" +
"        from t_cancel t1\n" +
"        LEFT JOIN repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " cust ON t1.cust_id = cust.cust_id    --------4\n" +
"                join  wxjy.jy_region_address_rel dz on t1.cust_id=dz.cust_id\n" +
"       group by t1.corp_org_id,dz.jy_region_name,dz.jy_non_region_name,t1.tclx, cust.cust_type\n" +
"               ) tt\n" +
"where tt.corp_org_id = 3328\n" +
" GROUP BY tt.corp_org_id,tt.区域,tt.非行政区域,tt.tclx,tt.cust_type";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "数字缴费等客户数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报商业缴费客户及用户数
        public static bool ybaobiao15()
        {
            string sqlString =
"select rad.jy_region_name 区域,\n" +
"       count(distinct scc.cust_id) 商业缴费客户数,\n" +
"       count(distinct sip.prod_inst_id) 商业缴费用户数\n" +
"  from so1.cm_customer               scc,\n" +
"       repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sip,\n" +
"       so1.ins_address               iad,wxjy.jy_region_address_rel rad\n" +
" where scc.cust_id = iad.cust_id(+)\n" +
"   and scc.cust_id = sip.cust_id\n" +
"   and scc.cust_type = 4    and scc.cust_id = rad.cust_id\n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and sip.prod_spec_id = 800200000001\n" +
"   and exists (select 1\n" +
"          from repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sis\n" +
"         where sis.prod_service_id = 1002\n" +
"           and sis.state = 1\n" +
"           and sis.os_status is null\n" +
"           and sis.prod_inst_id = sip.prod_inst_id)\n" +
" group by rad.jy_region_name";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "商业缴费客户及用户数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报宽带缴费、有效用户数
        public static bool ybaobiao16()
        {
            string sqlString =
"select rad.jy_region_name 区域,\n" +
" DECODE(scc.cust_type,\n" +
"              1,\n" +
"              '住宅客户',\n" +
"              4,\n" +
"              '普通非住宅客户',\n" +
"              7,\n" +
"              '集团非住宅客户',\n" +
"              '住宅客户') 客户类型,\n" +
"       count(distinct case\n" +
"               when (t1.state = '1' and t1.is_ins = '1' and t1.os_status is null or\n" +
"                    (t1.state = '99' or t1.is_ins = '0')) and\n" +
"                    pi.name not like '%测试%' and pi.name not like '%体验%' then\n" +
"                t1.prod_inst_id\n" +
"               else\n" +
"                null\n" +
"             end) 宽带缴费用户数,\n" +
"       count(distinct case\n" +
"               when (t1.state = '1' and t1.is_ins = '1' and t1.os_status is null or\n" +
"                    (t1.state = '99' or t1.is_ins = '0')) or\n" +
"                    (substr(t1.os_status, -1, 1) in ('1', '9') and\n" +
"                    t1.done_date > to_DATE('" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "', 'YYYYMMDD') - 365) or\n" +
"                    substr(t1.os_status, -1, 1) = '3' then\n" +
"                p.prod_inst_id\n" +
"               else\n" +
"                null\n" +
"             end) 宽带有效用户数\n" +
"  from repnew.fact_Ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "   p,\n" +
"       repnew.fact_Ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " t1,\n" +
"       product.up_product_item         pi,\n" +
"       so1.cm_customer                 scc,\n" +
"       wxjy.jy_region_address_rel      rad\n" +
" where scc.cust_id = rad.cust_id\n" +
"   and p.cust_id = scc.cust_id\n" +
"   and p.prod_inst_id = t1.prod_inst_id\n" +
"   and t1.prod_service_id = 1004\n" +
"   and p.corp_org_id = 3328\n" +
"   and t1.srvpkg_id = pi.product_item_id\n" +
" group by rad.jy_region_name, DECODE(scc.cust_type,\n" +
"              1,\n" +
"              '住宅客户',\n" +
"              4,\n" +
"              '普通非住宅客户',\n" +
"              7,\n" +
"              '集团非住宅客户',\n" +
"              '住宅客户')";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "宽带缴费、有效用户数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报集团缴费客户及用户数
        public static bool ybaobiao17()
        {
            string sqlString =
"select trim(case\n" +
"              when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 13, 3)\n" +
"              when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 6, 3)\n" +
"              when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"               '澄江'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"               '华士'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"               '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"               '祝塘'\n" +
"              when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"               '其他'\n" +
"              else\n" +
"               substr(scc.std_addr_name, 1, 3)\n" +
"            end) 区域,\n" +
"       count(distinct scc.cust_id) 客户数,\n" +
"       count(distinct sip.prod_inst_id) 用户数\n" +
"  from repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " scc, repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sip\n" +
" where scc.cust_id = sip.cust_id\n" +
"   and scc.cust_type = 7\n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and exists (select 1\n" +
"          from repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sis\n" +
"         where sis.prod_service_id = 1002\n" +
"           and sis.state = 1\n" +
"           and sis.os_status is null\n" +
"           and sis.prod_inst_id = sip.prod_inst_id)\n" +
" group by trim(case\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 13, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 6, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                  '澄江'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                  '华士'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                  '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                  '祝塘'\n" +
"                 when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"                  '其他'\n" +
"                 else\n" +
"                  substr(scc.std_addr_name, 1, 3)\n" +
"               end)";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "集团缴费客户及用户数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报高清机顶盒换机明细
        public static bool ybaobiao18()
        {
            string sqlString =
"select scc.cust_code 客户证号,scc.cust_name 姓名,rad.jy_region_name 区域,rrt1.serial_no 旧机顶盒号,rcd1.res_name 旧机顶盒类型,rrt2.serial_no 新机顶盒号,rcd2.res_name 新机顶盒类型\n" +
"  from so1.cm_customer               scc,\n" +
"       repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).AddDays(-1).ToString("yyyyMMdd") + " sip1,\n" +
"       res.res_terminal              rrt1,\n" +
"       res.res_code_definition       rcd1,\n" +
"       repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sip2,\n" +
"       res.res_terminal              rrt2,\n" +
"       res.res_code_definition       rcd2,\n" +
"       wxjy.jy_region_address_rel    rad\n" +
" where scc.cust_id = sip1.cust_id\n" +
"   and sip1.bill_id = rrt1.serial_no\n" +
"   and rrt1.res_code = rcd1.res_code\n" +
"   and scc.cust_id = sip2.cust_id\n" +
"   and sip1.prod_inst_id = sip2.prod_inst_id  and rcd1.res_type=2  and rcd2.res_type=2\n" +
"   and sip2.bill_id = rrt2.serial_no\n" +
"   and rrt2.res_code = rcd2.res_code\n" +
"   and scc.cust_id = rad.cust_id        and scc.cust_name not like '%测试%'  and scc.cust_prop<> 4\n" +
"   and rcd1.res_name not like '%高清%'\n" +
"   and rcd1.res_name not like '%4K%'\n" +
"   and (rcd2.res_name like '%高清%' or rcd2.res_name like '%4K%')";
            int[] columntxt = { 1,4,6 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "高清机顶盒换机明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报现金续费返充明细（360、720、返点）
        public static bool ybaobiao19()
        {
            string sqlString =
"select   rad.jy_region_name 区域,\n" +
"scc.cust_code 客户证号,\n" +
"       scc.cust_name 客户姓名,\n" +
"       ppd.promo_name 活动,\n" +
"       zpa.so_date 订购时间,\n" +
"       zpa.start_allot_bcycle 开始返充账期,\n" +
"       zpa.primary_fee / 100 返充总额,\n" +
"       zpa.alloted_fee / 100 已返充金额,\n" +
"       (zpa.primary_fee - zpa.alloted_fee) / 100 剩余返充金额\n" +
"  from zg.acc_pa_allot_info zpa, wxjy.jy_region_address_rel rad,\n" +
"       zg.acct              zac,\n" +
"       so1.cm_customer      scc,\n" +
"        so1.ins_address          iad,\n" +
"       zg.acc_pa_promo_def  ppd\n" +
" where zpa.acct_id = zac.acct_id\n" +
" and scc.cust_id = iad.cust_id    and scc.cust_id = rad.cust_id\n" +
"   and zac.cust_id = scc.cust_id\n" +
"   and zpa.promo_id = ppd.promo_id\n" +
"   and (ppd.promo_name like '%预存360元分2年返还套餐%'\n" +
"   or ppd.promo_name like '%高清互动升级充值360元分5年返还套餐%'\n" +
"   or ppd.promo_name like '%720元3年促销%'\n" +
"   or ppd.promo_name like '%720元5年促销%'\n" +
"   or ppd.promo_name like '%宽带猫押金返点%' \n" +
"   or ppd.promo_name like '%设备押金返点%') \n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and scc.cust_prop <> 4\n" +
"   and zpa.so_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"   and zpa.so_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
" order by scc.cust_code";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "现金续费返充明细（360、720、返点）--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报订购炫力动漫销账
        public static bool ybaobiao20()
        {
            string sqlString =
"select scc.cust_code 客户证号,\n" +
"       scc.cust_name 客户姓名,\n" +
"       rad.jy_region_name 区域,\n" +
"       iad.std_addr_name 地址,\n" +
"       nvl(aa.czje, 0) 实际出账金额_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + ",\n" +
"       nvl(bb.xzjine, 0) 销账金额_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "\n" +
"  from so1.cm_customer scc,\n" +
"       wxjy.jy_region_address_rel rad,\n" +
"       so1.ins_address iad,\n" +
"       (select cust.cust_code,\n" +
"               sum(cz.original_amount + cz.discount_amount +\n" +
"                   cz.adjust_amount) / 100 czje\n" +
"          from repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "    cust,\n" +
"               zg.acct                      zac,\n" +
"               repnew.fact_acct_item_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " cz\n" +
"         where cust.cust_id = zac.cust_id\n" +
"           and zac.acct_id = cz.acct_id\n" +
"           and cz.acct_item_type_id = 31276\n" +
"           and cust.own_corp_org_id = 3328\n" +
"         group by cust.cust_code) aa,\n" +
"       (select cust.cust_code, sum(xz.ppy_amount) / 100 xzjine\n" +
"          from repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "         cust,\n" +
"               zg.acct                           zac,\n" +
"               repnew.fact_payoff_dtl_new_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " xz\n" +
"         where cust.cust_id = zac.cust_id\n" +
"           and zac.acct_id = xz.acct_id\n" +
"           and xz.acct_item_type_id = 31276\n" +
"           and cust.own_corp_org_id = 3328\n" +
"         group by cust.cust_code) bb\n" +
" where scc.cust_code = aa.cust_code(+)\n" +
"   and scc.cust_code = bb.cust_code(+)\n" +
"   and scc.cust_id = rad.cust_id\n" +
"   and scc.cust_id = iad.cust_id\n" +
"   and (aa.czje <> 0 or bb.xzjine <> 0 )\n" +
" order by rad.jy_region_name, iad.std_addr_name, scc.cust_code";
            int[] columntxt = { 1 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "订购炫力动漫销账--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报充值卡销售详单
        public static bool ybaobiao21()
        {
            string sqlString =
           "select a.amount / 100 充值金额,\n" +
            "       cc.cust_code 客户证号,\n" +
            "       rad.jy_region_name 区域,\n" +
            "       p.card_no 充值卡号,\n" +
            "       a.payment_date 销售时间,\n" +
            "       bt.balance_type_name 账本\n" +
            "  from zg.acct_balance_log_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " a,\n" +
            "       zg.payment_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "          p,\n" +
            "       zg.acct                    c,\n" +
            "       so1.cm_customer            cc,\n" +
            "       zg.acct_balance_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "     t,\n" +
            "       zg.balance_type            bt,\n" +
            "       wxjy.jy_region_address_rel rad\n" +
            " where cc.cust_id = c.cust_id\n" +
            "   and a.payment_id = p.payment_id\n" +
            "   and p.acct_id = c.acct_id\n" +
            "   and cc.cust_id = rad.cust_id\n" +
            "   and a.operation_type = 103000\n" +
            "   and p.card_no is not null\n" +
            "   and c.corp_org_id = 3328\n" +
            "   and a.ACCT_BALANCE_ID = t.ACCT_BALANCE_ID\n" +
            "   and t.balance_type_id = bt.balance_type_id";
            int[] columntxt = { 2,4 }; //哪些列是文本格式
            int[] columndate = { 5 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "充值卡销售详单--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报各业务及账本销账金额
        public static bool ybaobiao22()
        {
            string sqlString =

"SELECT\n" +
"    decode(service_id, 1002,'数字基本', 1003, '互动基本产品', 1004, '宽带业务', 1005, '付费节目业务', 1006, '互动点播业务', 1008,'省互动增值') 分业务, t.区域, t.非行政区域,\n" +
"         t.cust_type,\n" +
"             SUM(sjcz_fee) 实际出账,\n" +
"         SUM(szjbxj_xz) 数字基本现金_销账,\n" +
"         SUM(kdxj_xz) 宽带现金_销账,\n" +
"         SUM(dbcz_xz) 点播充值卡_销账,\n" +
"         sum(xjty_xz) 现金账本_销账,\n" +
"         sum(xnty_xz) 虚拟通用_销账\n" +
"    from\n" +
"    (\n" +
"      SELECT\n" +
"     ais.service_id, bd.jy_non_region_name 非行政区域,bd.jy_region_name 区域,\n" +
"             DECODE(cust.cust_type,1,'住宅客户',4, '普通非住宅客户',7,'集团非住宅客户','住宅客户') cust_type,\n" +
"               SUM(ai.original_amount)/100 original_amount,\n" +
"               SUM(ai.adjust_amount)/100 adjust_amount, --调账\n" +
"               SUM(ai.discount_amount)/100 discount_amount, --优惠\n" +
"               SUM(ai.original_amount + ai.discount_amount+ai.adjust_amount)/100 sjcz_fee,\n" +
"                0 szjbxj_xz, --数字基本现金销账\n" +
"               0 kdxj_xz, --宽带现金账本销账\n" +
"               0 dbcz_xz, --点播业务充值卡账本销账\n" +
"               0 xjty_xz, --现金账本销账\n" +
"               0 xnty_xz--虚拟通用销账\n" +
"\n" +
"        FROM repnew.fact_acct_item_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " ai---1\n" +
"        JOIN zg.acct_item_service ais ON ai.acct_item_type_id = ais.acct_item_type_id\n" +
"        JOIN zg.acct acct ON ai.acct_id = acct.acct_id\n" +
"        LEFT JOIN repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "  cust ON acct.cust_id = cust.cust_id---2\n" +
"        LEFT JOIN wxjy.jy_region_address_rel bd ON cust.cust_id = bd.cust_id\n" +
"        WHERE ai.corp_org_id=3328\n" +
"       and cust.cust_prop <> 4\n" +
"        GROUP BY ais.service_id, bd.jy_non_region_name,bd.jy_region_name,cust.cust_type\n" +
"        UNION ALL\n" +
"        Select --cust.own_district_id own_district_id,\n" +
"        ais.service_id,bd.jy_non_region_name 非行政区域,bd.jy_region_name 区域,\n" +
"               DECODE(cust.cust_type,1,'住宅客户',4, '普通非住宅客户',7,'集团非住宅客户','住宅客户') cust_type,\n" +
"               0 original_amount,\n" +
"               0 adjust_amount, --调账\n" +
"               0 discount_amount, --优惠\n" +
"               0 sjcz_fee,\n" +
"              SUM(CASE WHEN pf.balance_type_id = 14 THEN pf.ppy_amount ELSE 0 END )/100 szjbxj_xz, --数字基本现金销账\n" +
"               SUM(CASE WHEN pf.balance_type_id = 22 THEN pf.ppy_amount ELSE 0 END )/100 kdxj_xz, --宽带现金账本销账\n" +
"               SUM(CASE WHEN pf.balance_type_id = 17 THEN pf.ppy_amount ELSE 0 END )/100 dbcz_xz, --点播业务充值卡账本销账\n" +
"               SUM(CASE WHEN pf.balance_type_id = 1 THEN pf.ppy_amount ELSE 0 END )/100 xjty_xz, --现金账本销账\n" +
"               SUM(CASE WHEN pf.balance_type_id = 3 THEN pf.ppy_amount ELSE 0 END )/100 xnty_xz --虚拟通用销账\n" +
"        FROM\n" +
"        repnew.fact_payoff_dtl_new_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " pf----3\n" +
"        JOIN zg.acct_item_service ais ON pf.acct_item_type_id = ais.acct_item_type_id\n" +
"        JOIN zg.acct acct ON pf.acct_id = acct.acct_id\n" +
"        LEFT JOIN repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " cust ON acct.cust_id = cust.cust_id----4\n" +
"        LEFT JOIN wxjy.jy_region_address_rel bd ON cust.cust_id = bd.cust_id\n" +
"        WHERE pf.corp_org_id=3328\n" +
"        and cust.cust_prop <> 4\n" +
"        GROUP By ais.service_id,bd.jy_non_region_name,bd.jy_region_name\n" +
"        ,cust.cust_type\n" +
"     ) t\n" +
"group by service_id,t.非行政区域,t.区域,\n" +
"         t.cust_type";

            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "各业务及账本销账金额--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报各业务账本余额
        public static bool ybaobiao23()
        {
            string sqlString =
"select ts.*,\n" +
"       ppy.ppyamt 批销,\n" +
"       pre_balance_amt + nvl(ppy.ppyamt, 0) 销账后余额\n" +
"  from ( select rel.jy_region_name 区域, rel.jy_non_region_name 非行政区域,\n" +
"               bt.balance_type_id,\n" +
"               bt.balance_type_name,\n" +
"               sum(t.balance) / 100 pre_balance_amt\n" +
"          from zg.acct_balance_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " t,\n" +
"               zg.acct                ac,\n" +
"               so1.cm_customer        scc,\n" +
"               so1.ins_address        iad,\n" +
"               zg.balance_type        bt,\n" +
"               wxjy.jy_region_address_rel rel\n" +
"         where t.corp_org_id = 3328\n" +
"           and t.acct_id = ac.acct_id\n" +
"           and ac.cust_id = scc.cust_id\n" +
"           and scc.cust_id = iad.cust_id(+)\n" +
"           and scc.cust_id = rel.cust_id\n" +
"           and t.balance_type_id = bt.balance_type_id\n" +
"         group by rel.jy_region_name, rel.jy_non_region_name,\n" +
"                  bt.balance_type_id,\n" +
"                  bt.balance_type_name) ts,\n" +
"       (select rel.jy_region_name 区域, rel.jy_non_region_name 非行政区域,\n" +
"               tt.balance_type_id,\n" +
"               bt.balance_type_name,\n" +
"               -sum(tt.ppy_amount) / 100 ppyamt\n" +
"          from repnew.fact_payoff_dtl_new_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " tt, ---临时表\n" +
"               zg.acct                           ac,\n" +
"               so1.cm_customer                   scc,\n" +
"               so1.ins_address                   iad,\n" +
"               zg.balance_type                   bt,\n" +
"               wxjy.jy_region_address_rel rel\n" +
"         where tt.corp_org_id = 3328\n" +
"           and tt.payment_date >= date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"           and tt.payment_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "' + 6/24\n" +
"           and tt.acct_id = ac.acct_id\n" +
"           and ac.cust_id = scc.cust_id\n" +
"           and scc.cust_id = iad.cust_id(+)\n" +
"           and scc.cust_id = rel.cust_id\n" +
"           and tt.balance_type_id = bt.balance_type_id\n" +
"         group by  rel.jy_region_name, rel.jy_non_region_name,\n" +
"                  tt.balance_type_id,\n" +
"                  bt.balance_type_name) ppy\n" +
" where ts.balance_type_id = ppy.balance_type_id(+)\n" +
"   and ts.非行政区域 = ppy.非行政区域(+)\n" +
"   and ts.区域 = ppy.区域(+)\n" +
" order by ts.区域, ts.非行政区域, ts.balance_type_id";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "各业务账本余额--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报新增机顶盒明细
        public static bool ybaobiao24()
        {
            string sqlString =
"select distinct scc.cust_code 客户证号,\n" +
"                scc.cust_name 客户姓名,\n" +
"                decode(scc.cust_type,\n" +
"                       '1',\n" +
"                       '公众客户',\n" +
"                       '4',\n" +
"                       '普通商业客户',\n" +
"                       '7',\n" +
"                       '合同商业客户') 客户类型,\n" +
"                ipr.res_equ_no 机顶盒号码,\n" +
"                sip.sub_bill_id 智能卡号码,\n" +
"                rcd.res_name 机顶盒类型,\n" +
"                con.cont_phone1 || ',' || con.cont_phone2 联系方式1,\n" +
"                con.cont_mobile1 || ',' || con.cont_mobile2 联系方式2,\n" +
"                rad.jy_region_name 区域,\n" +
"                iad.std_addr_name 地址\n" +
"  from so1.cm_customer            scc,\n" +
"       so1.ins_address            iad,\n" +
"       so1.cm_cust_contact_info   con,\n" +
"       so1.ins_prod               sip,\n" +
"       so1.ins_prod_res           ipr,\n" +
"       res.res_terminal           rrt,\n" +
"       res.res_code_definition    rcd,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where scc.cust_id = iad.cust_id\n" +
"   and scc.cust_id = rad.cust_id\n" +
"   and ipr.prod_inst_id = sip.prod_inst_id\n" +
"   and scc.cust_id = con.cust_id\n" +
"   and scc.cust_id = sip.cust_id\n" +
"   and rrt.res_code = rcd.res_code\n" +
"   and ipr.res_equ_no = rrt.serial_no\n" +
"   and rcd.res_type = 2\n" +
"   and ipr.create_date >= date '2020-" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("MM") + "-21'\n" +
"   and ipr.create_date < date '2020-" + DateTime.Now.ToString("MM") + "-20'\n" +
"   and scc.own_corp_org_id = 3328\n" +
" order by rad.jy_region_name, scc.cust_code, ipr.res_equ_no";
            int[] columntxt = { 1,4,5,7,8 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "新增机顶盒明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 月报新增机顶盒数量
        public static bool ybaobiao25()
        {
            string sqlString =
"select distinct rad.jy_region_name 区域,\n" +
"              count(distinct ipr.res_equ_no) 机顶盒数量\n" +
"  from so1.cm_customer            scc,\n" +
"       so1.ins_address            iad,\n" +
"       so1.cm_cust_contact_info   con,\n" +
"       so1.ins_prod               sip,\n" +
"       so1.ins_prod_res           ipr,\n" +
"       res.res_terminal           rrt,\n" +
"       res.res_code_definition    rcd,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where scc.cust_id = iad.cust_id\n" +
"   and scc.cust_id = rad.cust_id\n" +
"   and ipr.prod_inst_id = sip.prod_inst_id\n" +
"   and scc.cust_id = con.cust_id\n" +
"   and scc.cust_id = sip.cust_id\n" +
"   and rrt.res_code = rcd.res_code\n" +
"   and ipr.res_equ_no = rrt.serial_no\n" +
"   and rcd.res_type = 2\n" +
"   and ipr.create_date >= date '2020-" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("MM") + "-21'\n" +
"   and ipr.create_date < date '2020-" + DateTime.Now.ToString("MM") + "-20'\n" +
"   and scc.own_corp_org_id = 3328\n" +
"group by rad.jy_region_name";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "新增机顶盒数量--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion




        #region 各业务新增欠费金额
        public static bool ybaobiao26()
        {
            string sqlString =
                "select rad.jy_region_name 区域,\n" +
"       sum(case when ais.service_id = 1002 then zd.amount else 0 end) / 100 数字基本欠费金额,\n" +
"       sum(case when ais.service_id = 1003 then zd.amount else 0 end) / 100 互动基本产品欠费金额,\n" +
"       sum(case when ais.service_id = 1004 then zd.amount else 0 end) / 100 宽带业务欠费金额,\n" +
"       sum(case when ais.service_id = 1005 then zd.amount else 0 end) / 100 付费节目业务欠费金额,\n" +
"       sum(case when ais.service_id = 1006 then zd.amount else 0 end) / 100 互动点播业务欠费金额,\n" +
"       sum(case when ais.service_id = 1008 then zd.amount else 0 end) / 100 省互动增值欠费金额\n" +
"  from wxjy.jy_acct_item zd,\n" +
"       zg.acct zac,\n" +
"       so1.cm_customer scc,\n" +
"       zg.acct_item_service ais,\n" +
"       so1.ins_address iad,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where zd.acct_id = zac.acct_id\n" +
"   and zac.cust_id = scc.cust_id\n" +
"   and scc.cust_id = iad.cust_id\n" +
"   and scc.cust_id = rad.cust_id\n" +
"   and ais.acct_item_type_id = zd.acct_item_type_id\n" +
"   and zd.billing_cycle_id = 2020" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("MM") + "\n" +
"   and zd.amount > 0\n" +
"   and scc.own_corp_org_id = 3328\n" +
" group by rad.jy_region_name";

            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "各业务新增欠费金额--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion
        











        #region 经分出账客户数
        public static bool yjbaobiao1()
        {
            string sqlString =
"select trim(case\n" +
"              when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 13, 3)\n" +
"              when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 6, 3)\n" +
"              when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"               '澄江'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"               '华士'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"               '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"               '祝塘'\n" +
"              when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"               '其他'\n" +
"              else\n" +
"               substr(scc.std_addr_name, 1, 3)\n" +
"            end) 区域,\n" +
"       DECODE(scc.cust_type,\n" +
"              1,\n" +
"              '住宅客户',\n" +
"              4,\n" +
"              '普通非住宅客户',\n" +
"              7,\n" +
"              '集团非住宅客户',\n" +
"              '住宅客户') 客户类型,\n" +
"       count(distinct scc.cust_code) 出账客户数\n" +
"  from repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "    scc,\n" +
"       repnew.fact_acct_item_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " zd,\n" +
"       zg.acct                      zac\n" +
" where scc.cust_id = zac.cust_id\n" +
"   and zac.acct_id = zd.acct_id\n" +
"   and scc.own_corp_org_id = 3328\n" +
" group by trim(case\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 13, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 6, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                  '澄江'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                  '华士'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                  '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                  '祝塘'\n" +
"                 when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"                  '其他'\n" +
"                 else\n" +
"                  substr(scc.std_addr_name, 1, 3)\n" +
"               end),\n" +
"          DECODE(scc.cust_type,\n" +
"                 1,\n" +
"                 '住宅客户',\n" +
"                 4,\n" +
"                 '普通非住宅客户',\n" +
"                 7,\n" +
"                 '集团非住宅客户',\n" +
"                 '住宅客户')\n" +
"having SUM(zd.original_amount + zd.discount_amount + zd.adjust_amount) > 0\n" +
" order by trim(case\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 13, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 6, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                  '澄江'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                  '华士'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                  '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                  '祝塘'\n" +
"                 when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"                  '其他'\n" +
"                 else\n" +
"                  substr(scc.std_addr_name, 1, 3)\n" +
"               end)";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "出账客户数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 经分付费产品各站各分业务总量明细
        public static bool yjbaobiao2()
        {
            string sqlString =
"select trim(case\n" +
"              when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 13, 3)\n" +
"              when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 6, 3)\n" +
"              when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"               '澄江'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"               '华士'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"               '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"               '祝塘'\n" +
"              when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"               '其他'\n" +
"              else\n" +
"               substr(scc.std_addr_name, 1, 3)\n" +
"            end) 区域,\n" +
"       decode(sis.prod_service_id,\n" +
"              1002,\n" +
"              '数字基本',\n" +
"              1003,\n" +
"              '互动基本产品',\n" +
"              1004,\n" +
"              '宽带业务',\n" +
"              1005,\n" +
"              '付费节目业务',\n" +
"              1006,\n" +
"              '互动点播业务',\n" +
"              1008,\n" +
"              '省互动增值',\n" +
"              '') 分业务,\n" +
"       count(distinct sis.prod_inst_id) 订购总数\n" +
"  from repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sis,\n" +
"       repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "  sip,\n" +
"       repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "       scc,\n" +
"       product.up_product_item         upd\n" +
" where scc.cust_id = sip.cust_id\n" +
"   and sip.prod_inst_id = sis.prod_inst_id\n" +
"   and upd.product_item_id = sis.srvpkg_id\n" +
"   and sis.state = 1\n" +
"   and sis.os_status is null\n" +
"   and scc.own_corp_org_id = 3328\n" +
" group by trim(case\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 13, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 6, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                  '澄江'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                  '华士'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                  '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                  '祝塘'\n" +
"                 when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"                  '其他'\n" +
"                 else\n" +
"                  substr(scc.std_addr_name, 1, 3)\n" +
"               end),\n" +
"          sis.prod_service_id\n" +
" order by trim(case\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 13, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 6, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                  '澄江'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                  '华士'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                  '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                  '祝塘'\n" +
"                 when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"                  '其他'\n" +
"                 else\n" +
"                  substr(scc.std_addr_name, 1, 3)\n" +
"               end),\n" +
"          sis.prod_service_id";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "付费产品各站各分业务总量明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 经分付费各分业务各产品每月新增明细
        public static bool yjbaobiao3()
        {
            string sqlString =
"select upd.name 产品名称,\n" +
"       decode(sis.prod_service_id,\n" +
"              1002,\n" +
"              '数字基本',\n" +
"              1003,\n" +
"              '互动基本产品',\n" +
"              1004,\n" +
"              '宽带业务',\n" +
"              1005,\n" +
"              '付费节目业务',\n" +
"              1006,\n" +
"              '互动点播业务',\n" +
"              1008,\n" +
"              '省互动增值',\n" +
"              '') 分业务,\n" +
"       count(distinct sis.prod_inst_id) 新增订购数\n" +
"  from repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sis,\n" +
"       repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "   sip,\n" +
"       repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "      scc,\n" +
"       product.up_product_item         upd\n" +
" where scc.cust_id = sip.cust_id\n" +
"   and sip.prod_inst_id = sis.prod_inst_id\n" +
"   and upd.product_item_id = sis.srvpkg_id\n" +
"   and sis.state = 1\n" +
"   and sis.os_status is null\n" +
"   and sis.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and not exists (select 1\n" +
"          from repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).AddDays(-1).ToString("yyyyMMdd") + " sis1\n" +
"         where sis1.prod_inst_id = sip.prod_inst_id\n" +
"           and sis1.srvpkg_id = upd.product_item_id)\n" +
" group by upd.name, sis.prod_service_id\n" +
" order by upd.name, sis.prod_service_id";

            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "付费各分业务各产品每月新增明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 经分付费各消费等级客户数
        public static bool yjbaobiao4()
        {
            string sqlString =
"SELECT t.区域,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 0 and sjcz_fee <= 10) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额0到10元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 10 and sjcz_fee <= 20) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额10到20元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 20 and sjcz_fee <= 30) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额20到30元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 30 and sjcz_fee <= 40) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额30到40元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 40 and sjcz_fee <= 50) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额40到50元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 50 and sjcz_fee <= 100) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额50到100元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 100) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额大于100元\n" +
"  from (SELECT distinct scc.cust_code,\n" +
"                        trim(case\n" +
"                               when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                                    scc.std_addr_name not like '%测试%' then\n" +
"                                substr(scc.std_addr_name, 13, 3)\n" +
"                               when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                                    scc.std_addr_name not like '%测试%' then\n" +
"                                substr(scc.std_addr_name, 6, 3)\n" +
"                               when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                                '澄江'\n" +
"                               when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                                '华士'\n" +
"                               when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                                '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                               when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                                '祝塘'\n" +
"                               when (scc.std_addr_name like '%测试%' or\n" +
"                                    scc.std_addr_name is null) then\n" +
"                                '其他'\n" +
"                               else\n" +
"                                substr(scc.std_addr_name, 1, 3)\n" +
"                             end) 区域,\n" +
"                        SUM(ai.original_amount) / 100 original_amount,\n" +
"                        SUM(ai.adjust_amount) / 100 adjust_amount, --调账\n" +
"                        SUM(ai.discount_amount) / 100 discount_amount, --优惠\n" +
"                        SUM(ai.original_amount + ai.discount_amount +\n" +
"                            ai.adjust_amount) / 100 sjcz_fee\n" +
"          FROM repnew.fact_acct_item_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " ai ---1\n" +
"          JOIN zg.acct_item_service ais\n" +
"            ON ai.acct_item_type_id = ais.acct_item_type_id\n" +
"          JOIN zg.acct acct\n" +
"            ON ai.acct_id = acct.acct_id\n" +
"          LEFT JOIN repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " scc  ---2\n" +
"            ON acct.cust_id = scc.cust_id\n" +
"         WHERE ai.corp_org_id = 3328\n" +
"           and ais.service_id = 1005\n" +
"         GROUP BY trim(case\n" +
"                         when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                              scc.std_addr_name not like '%测试%' then\n" +
"                          substr(scc.std_addr_name, 13, 3)\n" +
"                         when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                              scc.std_addr_name not like '%测试%' then\n" +
"                          substr(scc.std_addr_name, 6, 3)\n" +
"                         when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                          '澄江'\n" +
"                         when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                          '华士'\n" +
"                         when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                          '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                         when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                          '祝塘'\n" +
"                         when (scc.std_addr_name like '%测试%' or\n" +
"                              scc.std_addr_name is null) then\n" +
"                          '其他'\n" +
"                         else\n" +
"                          substr(scc.std_addr_name, 1, 3)\n" +
"                       end),\n" +
"                  scc.cust_code) t\n" +
" group by t.区域";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "付费各消费等级客户数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 经分高清互动缴费客户、用户数
        public static bool yjbaobiao5()
        {
            string sqlString =
"select trim(case\n" +
"              when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 13, 3)\n" +
"              when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 6, 3)\n" +
"              when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"               '澄江'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"               '华士'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"               '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"               '祝塘'\n" +
"              when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"               '其他'\n" +
"              else\n" +
"               substr(scc.std_addr_name, 1, 3)\n" +
"            end) 区域,\n" +
"       DECODE(scc.cust_type,\n" +
"              1,\n" +
"              '住宅客户',\n" +
"              4,\n" +
"              '普通非住宅客户',\n" +
"              7,\n" +
"              '集团非住宅客户',\n" +
"              '住宅客户') 客户类型,\n" +
"       count(distinct scc.cust_code) 高清互动缴费客户数,\n" +
"       count(distinct sip.bill_id) 高清互动缴费用户数\n" +
"  from repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "  sip,\n" +
"       repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "     scc\n" +
" where sip.cust_id = scc.cust_id\n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and exists (select 1\n" +
"          from repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sis1\n" +
"         where sis1.prod_service_id = 1002\n" +
"           and sis1.state = 1\n" +
"           and sis1.os_status is null\n" +
"           and sis1.prod_inst_id = sip.prod_inst_id) --有正常基本包\n" +
"   and exists\n" +
" (select 1\n" +
"          from res.res_terminal rrt, res.res_code_definition rcd\n" +
"         where rrt.res_code = rcd.res_code\n" +
"           and (rcd.res_name like '%高清%' or rcd.res_name like '%4K%')\n" +
"           and sip.bill_id = rrt.serial_no) --有高清或者4K机顶盒\n" +
"and not exists\n" +
" (select 1\n" +
"          from res.res_terminal rrt, res.res_code_definition rcd\n" +
"         where rrt.res_code = rcd.res_code\n" +
"           and rcd.res_name in ('九州高清基本型DVC7058(江阴)','大亚高清交互型DC5000(江阴)','天柏高清基本型HMC0201BDH(江阴)','创维高清互动机顶盒_常熟','创维高清一体机(仪征)','九州高清交互型DVC7058EOC(江阴)','同洲高清交互型N8606(江阴)','同洲高清交互型5120de(江阴)')\n" +
"           and sip.bill_id = rrt.serial_no) --排除干扰项高清类型\n" +
"   and exists (select 1\n" +
"          from repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sis,\n" +
"               wxjy.rep_jy_itv_product         jip\n" +
"         where sis.srvpkg_id = jip.product_item_id\n" +
"           and sis.prod_inst_id = sip.prod_inst_id) --有互动产品\n" +
" group by trim(case\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 13, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 6, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                  '澄江'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                  '华士'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                  '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                  '祝塘'\n" +
"                 when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"                  '其他'\n" +
"                 else\n" +
"                  substr(scc.std_addr_name, 1, 3)\n" +
"               end),\n" +
"          DECODE(scc.cust_type,\n" +
"                 1,\n" +
"                 '住宅客户',\n" +
"                 4,\n" +
"                 '普通非住宅客户',\n" +
"                 7,\n" +
"                 '集团非住宅客户',\n" +
"                 '住宅客户')";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "高清互动缴费客户、用户数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 经分各站各分业务各产品每月新增明细
        public static bool yjbaobiao6()
        {
            string sqlString =
"select trim(case\n" +
"              when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 13, 3)\n" +
"              when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 6, 3)\n" +
"              when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"               '澄江'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"               '华士'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"               '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"               '祝塘'\n" +
"              when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"               '其他'\n" +
"              else\n" +
"               substr(scc.std_addr_name, 1, 3)\n" +
"            end) 区域,\n" +
"       upd.name 产品名称,\n" +
"       decode(sis.prod_service_id,\n" +
"              1002,\n" +
"              '数字基本',\n" +
"              1003,\n" +
"              '互动基本产品',\n" +
"              1004,\n" +
"              '宽带业务',\n" +
"              1005,\n" +
"              '付费节目业务',\n" +
"              1006,\n" +
"              '互动点播业务',\n" +
"              1008,\n" +
"              '省互动增值',\n" +
"              '') 分业务,\n" +
"       count(distinct sis.prod_inst_id) 新增订购数\n" +
"  from repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sis,\n" +
"       repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "   sip,\n" +
"       repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "       scc,\n" +
"       product.up_product_item         upd\n" +
" where scc.cust_id = sip.cust_id\n" +
"   and sip.prod_inst_id = sis.prod_inst_id\n" +
"   and upd.product_item_id = sis.srvpkg_id\n" +
"   and sis.state = 1\n" +
"   and sis.os_status is null\n" +
"   and sis.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and not exists (select 1\n" +
"          from repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).AddDays(-1).ToString("yyyyMMdd") + " sis1\n" +
"         where sis1.prod_inst_id = sip.prod_inst_id\n" +
"           and sis1.srvpkg_id = upd.product_item_id)\n" +
" group by trim(case\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 13, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 6, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                  '澄江'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                  '华士'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                  '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                  '祝塘'\n" +
"                 when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"                  '其他'\n" +
"                 else\n" +
"                  substr(scc.std_addr_name, 1, 3)\n" +
"               end),\n" +
"          upd.name,\n" +
"          sis.prod_service_id\n" +
" order by trim(case\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 13, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 6, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                  '澄江'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                  '华士'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                  '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                  '祝塘'\n" +
"                 when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"                  '其他'\n" +
"                 else\n" +
"                  substr(scc.std_addr_name, 1, 3)\n" +
"               end),\n" +
"          upd.name,\n" +
"          sis.prod_service_id";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "各站各分业务各产品每月新增明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 经分各站各分业务各产品总量明细
        public static bool yjbaobiao7()
        {
            string sqlString =
"select trim(case\n" +
"              when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 13, 3)\n" +
"              when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 6, 3)\n" +
"              when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"               '澄江'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"               '华士'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"               '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"               '祝塘'\n" +
"              when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"               '其他'\n" +
"              else\n" +
"               substr(scc.std_addr_name, 1, 3)\n" +
"            end) 区域,\n" +
"       upd.name 产品名称,\n" +
"       decode(sis.prod_service_id,\n" +
"              1002,\n" +
"              '数字基本',\n" +
"              1003,\n" +
"              '互动基本产品',\n" +
"              1004,\n" +
"              '宽带业务',\n" +
"              1005,\n" +
"              '付费节目业务',\n" +
"              1006,\n" +
"              '互动点播业务',\n" +
"              1008,\n" +
"              '省互动增值',\n" +
"              '') 分业务,\n" +
"       count(distinct sis.prod_inst_id) 总订购数\n" +
"  from repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sis,\n" +
"       repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "   sip,\n" +
"       repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "       scc,\n" +
"       product.up_product_item         upd\n" +
" where scc.cust_id = sip.cust_id\n" +
"   and sip.prod_inst_id = sis.prod_inst_id\n" +
"   and upd.product_item_id = sis.srvpkg_id\n" +
"   and sis.state = 1\n" +
"   and sis.os_status is null\n" +
"   and scc.own_corp_org_id = 3328\n" +
" group by trim(case\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 13, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 6, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                  '澄江'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                  '华士'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                  '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                  '祝塘'\n" +
"                 when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"                  '其他'\n" +
"                 else\n" +
"                  substr(scc.std_addr_name, 1, 3)\n" +
"               end),\n" +
"          upd.name,\n" +
"          sis.prod_service_id\n" +
" order by trim(case\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 13, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 6, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                  '澄江'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                  '华士'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                  '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                  '祝塘'\n" +
"                 when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"                  '其他'\n" +
"                 else\n" +
"                  substr(scc.std_addr_name, 1, 3)\n" +
"               end),\n" +
"          upd.name,\n" +
"          sis.prod_service_id";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "各站各分业务各产品总量明细--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 经分各账本的充值预收缴费汇总数据
        public static bool yjbaobiao8()
        {
            string sqlString =
"select /*+ PARALLEL(t,16) */\n" +
" rad.jy_region_name 区域,\n" +
" zbt.balance_type_name 账本名称,\n" +
" nvl(sum(case\n" +
"           when t.operation_type = 103000 and\n" +
"                (certified_type <> 100112 or certified_type is null) and\n" +
"                nvl(bank_id, 0) = 0 and card_no is null then\n" +
"            t.amount\n" +
"           else\n" +
"            null\n" +
"         end),\n" +
"     0) / 100 营业厅缴费,\n" +
" nvl(sum(case\n" +
"           when t.operation_type = 103000 and bank_id > 0 then\n" +
"            t.amount\n" +
"           else\n" +
"            null\n" +
"         end),\n" +
"     0) / 100 银行缴费,\n" +
" nvl(sum(case\n" +
"           when t.operation_type = 103000 and card_no is not null then\n" +
"            t.amount\n" +
"           else\n" +
"            null\n" +
"         end),\n" +
"     0) / 100 充值卡缴费,\n" +
" nvl(sum(case\n" +
"           when t.operation_type = 120000 then\n" +
"            t.amount\n" +
"           else\n" +
"            null\n" +
"         end),\n" +
"     0) / 100 批量预存,\n" +
" nvl(sum(case\n" +
"           when t.operation_type = 203000 then\n" +
"            t.amount\n" +
"           else\n" +
"            null\n" +
"         end),\n" +
"     0) / 100 预存反冲,\n" +
"\n" +
" nvl(sum(case\n" +
"           when t.operation_type = 110000 or\n" +
"                (certified_type = 100112 and t.operation_type = 103000) then\n" +
"            t.amount\n" +
"           else\n" +
"            null\n" +
"         end),\n" +
"     0) / 100 退费,\n" +
" nvl(sum(case\n" +
"           when t.operation_type = 141000 then\n" +
"            t.amount\n" +
"           else\n" +
"            null\n" +
"         end),\n" +
"     0) / 100 活动划拨预存,\n" +
" nvl(sum(case\n" +
"           when (t.operation_type = 103000 and\n" +
"                (certified_type <> 100112 or certified_type is null)) or\n" +
"                t.operation_type in (120000, 203000, 141000) or\n" +
"                (t.operation_type = 110000 or\n" +
"                (certified_type = 100112 and t.operation_type = 103000))\n" +
"            then\n" +
"            t.amount\n" +
"           else\n" +
"            null\n" +
"         end),\n" +
"     0) / 100 合计缴费\n" +
"  from (select acct.acct_id,\n" +
"               t.operation_type,\n" +
"               t.amount,\n" +
"               certified_type,\n" +
"               acct.corp_org_id,\n" +
"               t.balance_type_id,\n" +
"               b.card_no,\n" +
"               b.bank_id\n" +
"          from zg.acct_balance_log_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " t,\n" +
"               zg.payment_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + "          b,\n" +
"               zg.acct                    acct\n" +
"         where t.payment_id = b.payment_id(+)\n" +
"           and t.acct_id = acct.acct_id) t,\n" +
"       zg.balance_type zbt,\n" +
"       zg.acct zac,\n" +
"       so1.cm_customer scc,\n" +
"       so1.ins_address iad,\n" +
"       wxjy.jy_region_address_rel rad\n" +
" where t.balance_type_id = zbt.balance_type_id\n" +
"   and t.acct_id = zac.acct_id\n" +
"   and scc.cust_id = zac.cust_id\n" +
"   and scc.cust_id = iad.cust_id\n" +
"   and scc.cust_id = rad.cust_id\n" +
"   and scc.own_corp_org_id = 3328\n" +
" group by rad.jy_region_name, zbt.balance_type_name\n" +
" order by rad.jy_region_name, zbt.balance_type_name";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "各账本的充值预收缴费汇总数据--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 经分互动点播各消费等级客户数
        public static bool yjbaobiao9()
        {
            string sqlString =
"SELECT t.区域,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 0 and sjcz_fee <= 10) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额0到10元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 10 and sjcz_fee <= 20) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额10到20元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 20 and sjcz_fee <= 30) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额20到30元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 30 and sjcz_fee <= 40) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额30到40元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 40 and sjcz_fee <= 50) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额40到50元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 50 and sjcz_fee <= 100) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额50到100元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 100) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额大于100元\n" +
"  from (SELECT distinct scc.cust_code,\n" +
"                        trim(case\n" +
"                               when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                                    scc.std_addr_name not like '%测试%' then\n" +
"                                substr(scc.std_addr_name, 13, 3)\n" +
"                               when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                                    scc.std_addr_name not like '%测试%' then\n" +
"                                substr(scc.std_addr_name, 6, 3)\n" +
"                               when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                                '澄江'\n" +
"                               when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                                '华士'\n" +
"                               when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                                '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                               when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                                '祝塘'\n" +
"                               when (scc.std_addr_name like '%测试%' or\n" +
"                                    scc.std_addr_name is null) then\n" +
"                                '其他'\n" +
"                               else\n" +
"                                substr(scc.std_addr_name, 1, 3)\n" +
"                             end) 区域,\n" +
"                        SUM(ai.original_amount) / 100 original_amount,\n" +
"                        SUM(ai.adjust_amount) / 100 adjust_amount, --调账\n" +
"                        SUM(ai.discount_amount) / 100 discount_amount, --优惠\n" +
"                        SUM(ai.original_amount + ai.discount_amount +\n" +
"                            ai.adjust_amount) / 100 sjcz_fee\n" +
"          FROM repnew.fact_acct_item_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " ai ---1\n" +
"          JOIN zg.acct_item_service ais\n" +
"            ON ai.acct_item_type_id = ais.acct_item_type_id\n" +
"          JOIN zg.acct acct\n" +
"            ON ai.acct_id = acct.acct_id\n" +
"          LEFT JOIN repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " scc\n" +
"            ON acct.cust_id = scc.cust_id ---2\n" +
"         WHERE ai.corp_org_id = 3328\n" +
"           and ais.service_id = 1006\n" +
"         GROUP BY trim(case\n" +
"                         when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                              scc.std_addr_name not like '%测试%' then\n" +
"                          substr(scc.std_addr_name, 13, 3)\n" +
"                         when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                              scc.std_addr_name not like '%测试%' then\n" +
"                          substr(scc.std_addr_name, 6, 3)\n" +
"                         when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                          '澄江'\n" +
"                         when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                          '华士'\n" +
"                         when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                          '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                         when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                          '祝塘'\n" +
"                         when (scc.std_addr_name like '%测试%' or\n" +
"                              scc.std_addr_name is null) then\n" +
"                          '其他'\n" +
"                         else\n" +
"                          substr(scc.std_addr_name, 1, 3)\n" +
"                       end),\n" +
"                  scc.cust_code) t\n" +
" group by t.区域";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "互动点播各消费等级客户数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 经分互动基本的缴费客户数、缴费终端数
        public static bool yjbaobiao10()
        {
            string sqlString =
"select trim(case\n" +
"              when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 13, 3)\n" +
"              when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 6, 3)\n" +
"              when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"               '澄江'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"               '华士'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"               '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"               '祝塘'\n" +
"              when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"               '其他'\n" +
"              else\n" +
"               substr(scc.std_addr_name, 1, 3)\n" +
"            end) 区域,\n" +
"       DECODE(scc.cust_type,\n" +
"              1,\n" +
"              '住宅客户',\n" +
"              4,\n" +
"              '普通非住宅客户',\n" +
"              7,\n" +
"              '集团非住宅客户',\n" +
"              '住宅客户') 客户类型,\n" +
"       count(distinct scc.cust_id) 江阴互动基本缴费客户数,\n" +
"       count(distinct sip.prod_inst_id) 江阴互动基本缴费用户数\n" +
"  from repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sip, repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " scc\n" +
" where sip.cust_id = scc.cust_id\n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and exists (select 1\n" +
"          from repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sis\n" +
"         where sis.prod_service_id = 1002\n" +
"           and sis.state = 1\n" +
"           and sis.os_status is null\n" +
"           and sip.prod_inst_id = sis.prod_inst_id)\n" +
"   and exists\n" +
" (select 1\n" +
"          from repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sis1\n" +
"         where exists (select 1\n" +
"                  from product.up_product_item upd1\n" +
"                 where upd1.name = '江阴本地回看'\n" +
"                   and upd1.product_item_id = sis1.srvpkg_id)\n" +
"           and sip.prod_inst_id = sis1.prod_inst_id)\n" +
"   and exists\n" +
" (select 1\n" +
"          from repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sis2\n" +
"         where exists (select 1\n" +
"                  from product.up_product_item upd2\n" +
"                 where upd2.name = '栏目回看'\n" +
"                   and upd2.product_item_id = sis2.srvpkg_id)\n" +
"           and sip.prod_inst_id = sis2.prod_inst_id)\n" +
"   and exists\n" +
" (select 1\n" +
"          from repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sis3\n" +
"         where exists (select 1\n" +
"                  from product.up_product_item upd3\n" +
"                 where upd3.name = '频道回看'\n" +
"                   and upd3.product_item_id = sis3.srvpkg_id)\n" +
"           and sip.prod_inst_id = sis3.prod_inst_id)\n" +
"   and exists\n" +
" (select 1\n" +
"          from repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sis4\n" +
"         where exists (select 1\n" +
"                  from product.up_product_item upd4\n" +
"                 where upd4.name = '时移'\n" +
"                   and upd4.product_item_id = sis4.srvpkg_id)\n" +
"           and sip.prod_inst_id = sis4.prod_inst_id)\n" +
" group by trim(case\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 13, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 6, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                  '澄江'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                  '华士'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                  '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                  '祝塘'\n" +
"                 when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"                  '其他'\n" +
"                 else\n" +
"                  substr(scc.std_addr_name, 1, 3)\n" +
"               end),\n" +
"          DECODE(scc.cust_type,\n" +
"                 1,\n" +
"                 '住宅客户',\n" +
"                 4,\n" +
"                 '普通非住宅客户',\n" +
"                 7,\n" +
"                 '集团非住宅客户',\n" +
"                 '住宅客户')";

            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "互动基本的缴费客户数、缴费终端数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 经分互动基本各消费等级客户数
        public static bool yjbaobiao11()
        {
            string sqlString =
"SELECT t.区域,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 0 and sjcz_fee <= 10) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额0到10元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 10 and sjcz_fee <= 20) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额10到20元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 20 and sjcz_fee <= 30) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额20到30元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 30 and sjcz_fee <= 40) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额30到40元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 40 and sjcz_fee <= 50) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额40到50元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 50 and sjcz_fee <= 100) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额50到100元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 100) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额大于100元\n" +
"  from (SELECT distinct scc.cust_code,\n" +
"                        trim(case\n" +
"                               when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                                    scc.std_addr_name not like '%测试%' then\n" +
"                                substr(scc.std_addr_name, 13, 3)\n" +
"                               when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                                    scc.std_addr_name not like '%测试%' then\n" +
"                                substr(scc.std_addr_name, 6, 3)\n" +
"                               when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                                '澄江'\n" +
"                               when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                                '华士'\n" +
"                               when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                                '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                               when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                                '祝塘'\n" +
"                               when (scc.std_addr_name like '%测试%' or\n" +
"                                    scc.std_addr_name is null) then\n" +
"                                '其他'\n" +
"                               else\n" +
"                                substr(scc.std_addr_name, 1, 3)\n" +
"                             end) 区域,\n" +
"                        SUM(ai.original_amount) / 100 original_amount,\n" +
"                        SUM(ai.adjust_amount) / 100 adjust_amount, --调账\n" +
"                        SUM(ai.discount_amount) / 100 discount_amount, --优惠\n" +
"                        SUM(ai.original_amount + ai.discount_amount +\n" +
"                            ai.adjust_amount) / 100 sjcz_fee\n" +
"          FROM repnew.fact_acct_item_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " ai ---1\n" +
"          JOIN zg.acct_item_service ais\n" +
"            ON ai.acct_item_type_id = ais.acct_item_type_id\n" +
"          JOIN zg.acct acct\n" +
"            ON ai.acct_id = acct.acct_id\n" +
"          LEFT JOIN repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " scc  ---2\n" +
"            ON acct.cust_id = scc.cust_id\n" +
"         WHERE ai.corp_org_id = 3328\n" +
"           and ais.service_id = 1006\n" +
"         GROUP BY trim(case\n" +
"                         when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                              scc.std_addr_name not like '%测试%' then\n" +
"                          substr(scc.std_addr_name, 13, 3)\n" +
"                         when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                              scc.std_addr_name not like '%测试%' then\n" +
"                          substr(scc.std_addr_name, 6, 3)\n" +
"                         when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                          '澄江'\n" +
"                         when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                          '华士'\n" +
"                         when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                          '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                         when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                          '祝塘'\n" +
"                         when (scc.std_addr_name like '%测试%' or\n" +
"                              scc.std_addr_name is null) then\n" +
"                          '其他'\n" +
"                         else\n" +
"                          substr(scc.std_addr_name, 1, 3)\n" +
"                       end),\n" +
"                  scc.cust_code) t\n" +
" group by t.区域";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "互动基本各消费等级客户数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 经分宽带各消费等级客户数
        public static bool yjbaobiao12()
        {
            string sqlString =

"SELECT t.区域,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 0 and sjcz_fee <= 10) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额0到10元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 10 and sjcz_fee <= 20) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额10到20元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 20 and sjcz_fee <= 30) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额20到30元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 30 and sjcz_fee <= 40) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额30到40元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 40 and sjcz_fee <= 50) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额40到50元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 50 and sjcz_fee <= 100) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额50到100元,\n" +
"       count(distinct case\n" +
"               when (sjcz_fee > 100) then\n" +
"                cust_code\n" +
"               else\n" +
"                null\n" +
"             end) 消费金额大于100元\n" +
"  from (SELECT distinct scc.cust_code,\n" +
"                        trim(case\n" +
"                               when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                                    scc.std_addr_name not like '%测试%' then\n" +
"                                substr(scc.std_addr_name, 13, 3)\n" +
"                               when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                                    scc.std_addr_name not like '%测试%' then\n" +
"                                substr(scc.std_addr_name, 6, 3)\n" +
"                               when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                                '澄江'\n" +
"                               when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                                '华士'\n" +
"                               when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                                '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                               when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                                '祝塘'\n" +
"                               when (scc.std_addr_name like '%测试%' or\n" +
"                                    scc.std_addr_name is null) then\n" +
"                                '其他'\n" +
"                               else\n" +
"                                substr(scc.std_addr_name, 1, 3)\n" +
"                             end) 区域,\n" +
"                        SUM(ai.original_amount) / 100 original_amount,\n" +
"                        SUM(ai.adjust_amount) / 100 adjust_amount, --调账\n" +
"                        SUM(ai.discount_amount) / 100 discount_amount, --优惠\n" +
"                        SUM(ai.original_amount + ai.discount_amount +\n" +
"                            ai.adjust_amount) / 100 sjcz_fee\n" +
"          FROM repnew.fact_acct_item_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " ai ---1\n" +
"          JOIN zg.acct_item_service ais\n" +
"            ON ai.acct_item_type_id = ais.acct_item_type_id\n" +
"          JOIN zg.acct acct\n" +
"            ON ai.acct_id = acct.acct_id\n" +
"          LEFT JOIN repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " scc\n" +
"            ON acct.cust_id = scc.cust_id ---2\n" +
"         WHERE ai.corp_org_id = 3328\n" +
"           and ais.service_id = 1004\n" +
"         GROUP BY trim(case\n" +
"                         when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                              scc.std_addr_name not like '%测试%' then\n" +
"                          substr(scc.std_addr_name, 13, 3)\n" +
"                         when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                              scc.std_addr_name not like '%测试%' then\n" +
"                          substr(scc.std_addr_name, 6, 3)\n" +
"                         when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                          '澄江'\n" +
"                         when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                          '华士'\n" +
"                         when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                          '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                         when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                          '祝塘'\n" +
"                         when (scc.std_addr_name like '%测试%' or\n" +
"                              scc.std_addr_name is null) then\n" +
"                          '其他'\n" +
"                         else\n" +
"                          substr(scc.std_addr_name, 1, 3)\n" +
"                       end),\n" +
"                  scc.cust_code) t\n" +
" group by t.区域";

            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "宽带各消费等级客户数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 经分数字产品月报
        public static bool yjbaobiao13()
        {
            string sqlString =
"select  rad.jy_region_name 区域,\n" +
"       count(distinct  case\n" +
"             when  pup.name like '%幸福放映厅%' then\n" +
"                sip.prod_inst_id\n" +
"               else\n" +
"                null\n" +
"             end)  幸福放映厅,\n" +
"      count(distinct  case\n" +
"             when  pup.name like '%老少同乐包%' then\n" +
"                sip.prod_inst_id\n" +
"               else\n" +
"                null\n" +
"             end)  老少同乐包,\n" +
"               count(distinct  case\n" +
"             when  pup.name like '%幸福100%' then\n" +
"                sip.prod_inst_id\n" +
"               else\n" +
"                null\n" +
"             end)  幸福100,\n" +
"               count(distinct  case\n" +
"             when  pup.name like '%体育休闲%' then\n" +
"                sip.prod_inst_id\n" +
"               else\n" +
"                null\n" +
"             end)  体育休闲,\n" +
"               count(distinct  case\n" +
"             when  pup.name like '%nvod%' then\n" +
"                sip.prod_inst_id\n" +
"               else\n" +
"                null\n" +
"             end)  nvod,\n" +
"               count(distinct  case\n" +
"             when  pup.name like '%戏曲点播%' then\n" +
"                sip.prod_inst_id\n" +
"               else\n" +
"                null\n" +
"             end)  戏曲点播,\n" +
"             count(distinct  case\n" +
"             when  pup.name like '%HBO鼎级剧场%' then\n" +
"                sip.prod_inst_id\n" +
"               else\n" +
"                null\n" +
"             end)  HBO鼎级剧场,\n" +
"             count(distinct  case\n" +
"             when  pup.name like '%好莱坞专区%' then\n" +
"                sip.prod_inst_id\n" +
"               else\n" +
"                null\n" +
"             end)  好莱坞专区,\n" +
"             count(distinct  case\n" +
"             when  pup.name like '%全家福点播%' then\n" +
"                sip.prod_inst_id\n" +
"               else\n" +
"                null\n" +
"             end)  全家福点播,\n" +
"             count(distinct  case\n" +
"             when  pup.name like '%优加互动彩虹岛%' then\n" +
"                sip.prod_inst_id\n" +
"               else\n" +
"                null\n" +
"             end)  优加互动彩虹岛,\n" +
"             count(distinct  case\n" +
"             when  pup.name like '%优加互动广场舞%' then\n" +
"                sip.prod_inst_id\n" +
"               else\n" +
"                null\n" +
"             end)  优加互动广场舞\n" +
"  from product.up_product_item    pup,   wxjy.jy_region_address_rel rad,\n" +
"       so1.ins_srvpkg         sis,\n" +
"       so1.ins_prod             sip,\n" +
"       so1.cm_customer          scc,\n" +
"       so1.ins_address          iad\n" +
" where pup.product_item_id  =  sis.srvpkg_id   and scc.cust_id = rad.cust_id\n" +
" and sip.prod_inst_id = sis.prod_inst_id\n" +
"   and sip.cust_id = scc.cust_id\n" +
"   and scc.cust_id = iad.cust_id\n" +
"   and sis.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"   and sis.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"   and scc.own_corp_org_id = 3328\n" +
" group by rad.jy_region_name";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "数字产品月报--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 经分数字新增客户数
        public static bool yjbaobiao14()
        {
            string sqlString =
"select trim(case\n" +
"              when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 13, 3)\n" +
"              when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 6, 3)\n" +
"              when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"               '澄江'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"               '华士'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"               '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"               '祝塘'\n" +
"              when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"               '其他'\n" +
"              else\n" +
"               substr(scc.std_addr_name, 1, 3)\n" +
"            end) 区域,\n" +
"       count(distinct scc.cust_code) 新增客户数\n" +
"  from repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " scc\n" +
" where scc.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"   and scc.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"   and scc.cust_type = 1\n" +
"   and scc.own_corp_org_id = 3328\n" +
"   and exists (select 1\n" +
"          from repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "   sip,\n" +
"               repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sis\n" +
"         where sip.prod_inst_id = sis.prod_inst_id\n" +
"           and sis.prod_service_id = 1002\n" +
"           and sis.state = 1\n" +
"           and sis.os_status is null\n" +
"           and sip.cust_id = scc.cust_id)\n" +
" group by trim(case\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 13, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 6, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                  '澄江'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                  '华士'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                  '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                  '祝塘'\n" +
"                 when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"                  '其他'\n" +
"                 else\n" +
"                  substr(scc.std_addr_name, 1, 3)\n" +
"               end)";

            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "数字新增客户数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 经分数字电视新增用户数按区域统计
        public static bool yjbaobiao15()
        {
            string sqlString =
"select trim(case\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 13, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 6, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                  '澄江'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                  '华士'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                  '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                  '祝塘'\n" +
"                 when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"                  '其他'\n" +
"                 else\n" +
"                  substr(scc.std_addr_name, 1, 3)\n" +
"               end) 区域,\n" +
"       count(distinct op.prod_inst_id) 新增用户数\n" +
"  from (select *\n" +
"          from so1.ord_cust\n" +
"        union all\n" +
"        select *\n" +
"          from so1.ord_cust_f_2020) oc,\n" +
"       (select *\n" +
"          from so1.ord_prod\n" +
"        union all\n" +
"        select *\n" +
"          from so1.ord_prod_f_2020) op,\n" +
"       (select *\n" +
"          from so1.ord_srvpkg\n" +
"        union all\n" +
"        select *\n" +
"          from so1.ord_srvpkg_f_2020) os,\n" +
"       repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " scc\n" +
" where scc.cust_id = op.cust_id\n" +
"   and oc.CUST_ORDER_ID = op.CUST_ORDER_ID\n" +
"   and op.prod_order_id = os.prod_order_id\n" +
"   and oc.BUSINESS_ID in (800001000001, 800001000002) --普通新装,批量新装\n" +
"   and op.prod_spec_id = 800200000001 --宽带:800200000003   数字:800200000001\n" +
"   and oc.order_state <> 10 --排除新装撤单情况\n" +
"   and oc.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"   and oc.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
"   and oc.own_corp_org_id = 3328\n" +
" group by trim(case\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 13, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 6, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                  '澄江'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                  '华士'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                  '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                  '祝塘'\n" +
"                 when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"                  '其他'\n" +
"                 else\n" +
"                  substr(scc.std_addr_name, 1, 3)\n" +
"               end)";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "数字电视新增用户数按区域统计--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 经分数字套餐月报
        public static bool yjbaobiao16()
        {
            string sqlString =

"select rad.jy_region_name 区域,\n" +
"       count(distinct case\n" +
"               when iof.offer_id = 800056 then\n" +
"                sip.prod_inst_id\n" +
"               else\n" +
"                null\n" +
"             end) 订购720元分5年促销,\n" +
"       count(distinct case\n" +
"               when iof.offer_id = 800057 then\n" +
"                sip.prod_inst_id\n" +
"               else\n" +
"                null\n" +
"             end) 订购720元分3年促销,\n" +
"               count(distinct case\n" +
"               when iof.offer_id = 800500211162 then\n" +
"                sip.prod_inst_id\n" +
"               else\n" +
"                null\n" +
"             end) 好视乐炫互动套餐_新装,\n" +
"             count(distinct case\n" +
"               when iof.offer_id = 800500211163 then\n" +
"                sip.prod_inst_id\n" +
"               else\n" +
"                null\n" +
"             end) 好视乐炫宽带套餐_新装,\n" +
"             count(distinct case\n" +
"               when iof.offer_id = 800500211161 then\n" +
"                sip.prod_inst_id\n" +
"               else\n" +
"                null\n" +
"             end) 好视乐炫高清套餐_新装,\n" +
"              count(distinct case\n" +
"               when iof.offer_id = 800500012453 then\n" +
"                sip.prod_inst_id\n" +
"               else\n" +
"                null\n" +
"             end) 优加89_套餐,\n" +
"\n" +
"       count(distinct case\n" +
"               when iof.offer_id = 800500012448 then\n" +
"                sip.prod_inst_id\n" +
"               else\n" +
"                null\n" +
"             end) 优加69_套餐\n" +
"  from so1.ins_offer            iof,wxjy.jy_region_address_rel rad,\n" +
"       so1.ins_off_ins_prod_rel rel,\n" +
"       so1.ins_prod             sip,\n" +
"       so1.cm_customer          scc,\n" +
"       so1.ins_address          iad\n" +
" where iof.offer_inst_id = rel.offer_inst_id and scc.cust_id = rad.cust_id\n" +
"   and rel.prod_inst_id = sip.prod_inst_id\n" +
"   and sip.cust_id = scc.cust_id\n" +
"   and scc.cust_id = iad.cust_id\n" +
"   and iof.offer_id in\n" +
"       (800412, 800413, 800418, 800419, 800420, 800425, 800426, 800427,800500211162,800500211163,800500211161,800056,800057,800500012453,800500012448)\n" +
"   and iof.create_date >= date '" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyy-MM-dd") + "'\n" +
"   and iof.create_date < date '" + DateTime.Now.ToString("yyyy-MM-01") + "'\n" +
" group by rad.jy_region_name";

            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "数字套餐月报--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 经分数字有效用户数
        public static bool yjbaobiao17()
        {
            string sqlString =
"select trim(case\n" +
"              when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 13, 3)\n" +
"              when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 6, 3)\n" +
"              when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"               '澄江'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"               '华士'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"               '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"               '祝塘'\n" +
"              when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"               '其他'\n" +
"              else\n" +
"               substr(scc.std_addr_name, 1, 3)\n" +
"            end) 区域,\n" +
"       --dz.jy_non_region_name 非行政区域,\n" +
"       count(distinct sip.prod_inst_id) 数字有效用户数\n" +
"  from repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sip, repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " scc\n" +
" where sip.cust_id = scc.cust_id\n" +
"   and sip.is_dtv = 1 --数字电视\n" +
"   and sip.is_valid1 = 1 --有效\n" +
"   and scc.own_corp_org_id = 3328\n" +
" group by trim(case\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 13, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 6, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                  '澄江'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                  '华士'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                  '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                  '祝塘'\n" +
"                 when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"                  '其他'\n" +
"                 else\n" +
"                  substr(scc.std_addr_name, 1, 3)\n" +
"               end)\n" +
" order by trim(case\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 13, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 6, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                  '澄江'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                  '华士'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                  '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                  '祝塘'\n" +
"                 when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"                  '其他'\n" +
"                 else\n" +
"                  substr(scc.std_addr_name, 1, 3)\n" +
"               end)";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "数字有效用户数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 经分销账客户数
        public static bool yjbaobiao18()
        {
            string sqlString =

"select trim(case\n" +
"              when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 13, 3)\n" +
"              when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 6, 3)\n" +
"              when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"               '澄江'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"               '华士'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"               '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"               '祝塘'\n" +
"              when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"               '其他'\n" +
"              else\n" +
"               substr(scc.std_addr_name, 1, 3)\n" +
"            end) 区域,\n" +
"       DECODE(scc.cust_type,\n" +
"              1,\n" +
"              '住宅客户',\n" +
"              4,\n" +
"              '普通非住宅客户',\n" +
"              7,\n" +
"              '集团非住宅客户',\n" +
"              '住宅客户') 客户类型,\n" +
"       count(distinct scc.cust_code) 销账客户数\n" +
"  from repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "       scc,\n" +
"       repnew.fact_payoff_dtl_new_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMM") + " t,\n" +
"       zg.acct                           tt\n" +
" where t.acct_id = tt.acct_id\n" +
"   and tt.cust_id = scc.cust_id\n" +
"   and t.ppy_amount > 0\n" +
"   and scc.own_corp_org_id = 3328\n" +
" group by trim(case\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 13, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 6, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                  '澄江'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                  '华士'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                  '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                  '祝塘'\n" +
"                 when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"                  '其他'\n" +
"                 else\n" +
"                  substr(scc.std_addr_name, 1, 3)\n" +
"               end),\n" +
"          DECODE(scc.cust_type,\n" +
"                 1,\n" +
"                 '住宅客户',\n" +
"                 4,\n" +
"                 '普通非住宅客户',\n" +
"                 7,\n" +
"                 '集团非住宅客户',\n" +
"                 '住宅客户')";

            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "销账客户数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

        #region 经分自19年底以来的留存客户数
        public static bool yjbaobiao19()
        {
            string sqlString =
"select trim(case\n" +
"              when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 13, 3)\n" +
"              when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                   scc.std_addr_name not like '%测试%' then\n" +
"               substr(scc.std_addr_name, 6, 3)\n" +
"              when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"               '澄江'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"               '华士'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"               '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"               '祝塘'\n" +
"              when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"               '其他'\n" +
"              else\n" +
"               substr(scc.std_addr_name, 1, 3)\n" +
"            end) 区域,\n" +
"       decode(scc.cust_type,\n" +
"              1,\n" +
"              '公众客户',\n" +
"              4,\n" +
"              '普通商业客户',\n" +
"              7,\n" +
"              '合同商业客户') 客户类型,\n" +
"       count(distinct scc.cust_id) 留存缴费客户数\n" +
"  from repnew.fact_cust_" + DateTime.Parse(DateTime.Now.ToString("yyyy-01-01")).AddDays(-1).ToString("yyyyMMdd") + " scc\n" +
" where scc.own_corp_org_id = 3328\n" +
"   and exists (select 1\n" +
"          from repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-01-01")).AddDays(-1).ToString("yyyyMMdd") + "   sip,\n" +
"               repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-01-01")).AddDays(-1).ToString("yyyyMMdd") + " sis\n" +
"         where sis.prod_inst_id = sip.prod_inst_id\n" +
"           and sis.prod_service_id = 1002\n" +
"           and sis.state = 1\n" +
"           and sis.os_status is null\n" +
"           and sip.cust_id = scc.cust_id)\n" +
"   and exists (select 1\n" +
"          from repnew.fact_ins_prod_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + "   sip1,\n" +
"               repnew.fact_ins_srvpkg_" + DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd") + " sis1\n" +
"         where sis1.prod_inst_id = sip1.prod_inst_id\n" +
"           and sis1.prod_service_id = 1002\n" +
"           and sis1.state = 1\n" +
"           and sis1.os_status is null\n" +
"           and sip1.cust_id = scc.cust_id)\n" +
" group by trim(case\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '江苏省' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 13, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 4) = '江阴公司' and\n" +
"                      scc.std_addr_name not like '%测试%' then\n" +
"                  substr(scc.std_addr_name, 6, 3)\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '澄江站' then\n" +
"                  '澄江'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '华士华' then\n" +
"                  '华士'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '南闸南' then\n" +
"                  '南闸'\n" +
"              when substr(scc.std_addr_name, 1, 3) = '申港申' then\n" +
"               '申港'\n" +
"                 when substr(scc.std_addr_name, 1, 3) = '祝塘祝' then\n" +
"                  '祝塘'\n" +
"                 when (scc.std_addr_name like '%测试%' or scc.std_addr_name is null) then\n" +
"                  '其他'\n" +
"                 else\n" +
"                  substr(scc.std_addr_name, 1, 3)\n" +
"               end),\n" +
"          decode(scc.cust_type,\n" +
"                 1,\n" +
"                 '公众客户',\n" +
"                 4,\n" +
"                 '普通商业客户',\n" +
"                 7,\n" +
"                 '合同商业客户')";
            int[] columntxt = { 0 }; //哪些列是文本格式
            int[] columndate = { 0 };         //哪些列是日期格式
            DataTable dt = OracleHelper.ExecuteDataTable(sqlString);
            ExcelHelper.DataTableToExcel("\\月报\\" + "自19年底以来的留存客户数--" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx", dt, sqlString, columntxt, columndate);
            return true;
        }
        #endregion

    }
}
