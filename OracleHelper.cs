using Oracle.ManagedDataAccess.Client;
using System;
using System.Data;

namespace AutoUpDataBoss
{
    class OracleHelper
    {
        private static string connStr = "User Id=wxjy;Password=wxjy1234!;Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=172.28.240.218)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=cloudboss)))";

        #region 执行SQL语句,返回受影响行数
        public static bool ExecuteNonQuery(string sql/*, params OracleParameter[] parameters*/)
        {
            using (OracleConnection conn = new OracleConnection(connStr))
            {
                conn.Open();
                using (OracleCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = sql;
                    //  cmd.Parameters.AddRange(parameters);
                    cmd.ExecuteNonQuery();
                    conn.Dispose();//释放连接资源。
                    conn.Close();//关闭连接。
                    OracleConnection.ClearPool(conn);//彻底关闭链接。
                    return true;
                }
            }
        }
        #endregion
        #region 执行SQL语句,返回DataTable;只用来执行查询结果比较少的情况
        public static DataTable ExecuteDataTable(string sql/*, params OracleParameter[] parameters*/)
        {
            using (OracleConnection conn = new OracleConnection(connStr))
            {
                    conn.Open();
                    using (OracleCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = sql;
                        //  cmd.Parameters.AddRange(parameters);
                        OracleDataAdapter adapter = new OracleDataAdapter(cmd);
                        DataTable datatable = new DataTable();
                        adapter.Fill(datatable);
                        conn.Dispose();//释放连接资源。
                        conn.Close();//关闭连接。
                        OracleConnection.ClearPool(conn);//彻底关闭链接。
                    return datatable;
                    }
            }
        }
        #endregion

        //#region 执行SQL语句,返回DataTable;只用来执行查询结果比较少的情况
        //public void Dispose()
        //{
        //    if (conn != null)//当连接对象不为空时执行。
        //    {
        //        conn.Close();//关闭SQL连接。
        //        conn.Dispose();//释放SQL连接资源。
        //    }
        //    //  GC.Collect();//用完所占用的内存资源后，进行可能的垃圾回收以释放不再需要的大量内存（注：垃圾回收器可以确定最佳的垃圾回收时间，但不可多用）。
        //}
        //#endregion
    }
}
