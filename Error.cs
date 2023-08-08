using System;
using System.IO;//操作文件

namespace AutoUpDataBoss
{
    class Error
    {
        private StreamWriter swWrite = null;//写文件
        private string logpath = null;//日志路径



        /// <summary>
        /// 判断是否有文件目录，如果没有则创建
        /// </summary>
        /// <param name="pathweb">路径</param>
        private void InitDir(string pathweb)
        {
            if (!Directory.Exists(pathweb))
            {
                Directory.CreateDirectory(pathweb);
            }
        }

        /// <summary>
        /// 判断是否有当天日志，如果没有则创建
        /// </summary>
        /// <param name="sCurDate">日期</param>
        /// <param name="pathweb">路径</param>
        private void InitLog(string sCurDate, string pathweb)
        {
            if (sCurDate != "" || sCurDate != null)//判断日期值是否合法
            {
                logpath = pathweb + "/AutoDataError/" + sCurDate + ".txt";
            }
            if (!File.Exists(logpath))//判断今日的日志有没,没有就创建，有就继续添加
            {
                swWrite = File.CreateText(logpath);//创建
            }
            else
            {
                swWrite = File.AppendText(logpath);//已有的继续添加
            }

        }

        /// <summary>
        /// 日志主方法
        /// </summary>
        /// <param name="sMess">错误</param>
        /// <param name="ex">异常</param>
        /// <param name="path1">日志路径</param>
        public void LogError(string sMess, Exception ex)
        {
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            string path = pathInfo.Parent.FullName;
            DateTime dt = DateTime.Now;
            string NowYear = dt.ToString("yyyy");
            string NowMouth = dt.ToString("MM");
            string NowDay = dt.ToString("dd");
            string sCurDate = NowYear + "-" + NowMouth + "-" + NowDay;
            if (swWrite == null)
            {
                InitDir(path + "/AutoDataError");//调用文件创建方法
                InitLog(sCurDate, path);//调用日志创建方法
            }
            swWrite.WriteLine(dt.ToString("yyyy-MM-dd HH:mm:ss"));//错误产生日期
            swWrite.WriteLine("输出信息：" + sMess);
            if (ex != null)
            {
                swWrite.WriteLine("异常信息：\r\n" + ex.ToString());//创建异常信息
                swWrite.WriteLine("异常堆栈：\r\n" + ex.StackTrace);//创建异常堆栈
            }
            swWrite.WriteLine("\r\n");
            swWrite.Flush();
            swWrite.Dispose();
        }
    }
}
