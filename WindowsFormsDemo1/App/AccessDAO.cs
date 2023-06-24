using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Threading.Tasks;
using System.Text.RegularExpressions; //正则表达式引用所需

namespace WindowsFormsDemo1.App
{
    //access的数据访问接口
    class AccessDAO
    {
        public static class Property
        {
            public static string accessFilePath = @"D:\zhuomian\无人机\数据库设计\DatabaseDemo2.mdb";
            //若放入主程序，则可如下设置
            //one mainFrm = (one)this.Owner;
            //string prjName = mainFrm.laPrj.Text;
            //string prjPath = mainFrm.laFile_Path.Text;
            // public static string accessFilePath = prjPath + "\\矢量数据\\" + prjName + ".mdb";
        }

        //从access数据库获取数据
        //dataFilePath指定access文件的路径
        //sql指定数据库的查询语句
        //DataSet为查询返回的数据集
        public static DataSet getDataSetFromAccessTable(string sql, string dataFilePath)
        {
            // 连接数据库
            OleDbConnection connct = new OleDbConnection();
            string oleDB = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dataFilePath;
            connct.ConnectionString = oleDB;

            //创建命令
            OleDbCommand command = new OleDbCommand(sql, connct);

            //打开数据库
            connct.Open();

            //执行命令
            DataSet dataSet = new DataSet();
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(command);

            dataAdapter.Fill(dataSet);

            // 关闭连接
            connct.Close();
            return dataSet;
        }


        //更新或者插入数据到access数据库
        //dataFilePath指定access文件的路径
        //sql指定数据库的更新或者插入语句
        //返回值int表示此次更新影响的行数
        public static int updateAccessTable(string sql, string dataFilePath)
        {
            // 连接数据库
            OleDbConnection connct = new OleDbConnection();
            string oleDB = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dataFilePath;
            connct.ConnectionString = oleDB;

            //打开数据库
            connct.Open();

            //执行命令
            OleDbCommand myCommand = new OleDbCommand(sql, connct);
            int res = myCommand.ExecuteNonQuery();

            // 关闭连接
            connct.Close();
            return res;
        }

        //更新或者插入数据到access数据库
        //dataFilePath指定access文件的路径
        //command指定操作（更新或者插入）数据库的命令
        //返回值int表示此次更新影响的行数
        public static int updateAccessTable(OleDbCommand command, string dataFilePath)
        {
            // 连接数据库
            OleDbConnection connct = new OleDbConnection();
            string oleDB = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dataFilePath;
            connct.ConnectionString = oleDB;

            //打开数据库
            connct.Open();

            //执行命令
            //OleDbCommand myCommand = new OleDbCommand(sql, connct);
            command.Connection = connct;
            int res = command.ExecuteNonQuery();

            // 关闭连接
            connct.Close();
            return res;
        }



    }
}
