using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace EPPlusWeb.Models
{
    public class ExcelData
    {
        public static DataSet GetExcelData()
        {
            DataSet ds = new DataSet();
            string[,] infos =
            {
                { "151100310001","刘备","男","计算机科学与工程学院","计算机科学与技术"},
                { "151100310002","关羽","男","计算机科学与工程学院","通信工程"},
                { "151100310003","张飞","男","数学与统计学院","信息与计算科学"},
                { "151100310004","小乔","女","文学院","汉语言文学"}
            };
            string[,] scores =
            {
                { "151100310001","刘备","88","90","80"},
                { "151100310002","关羽","86","70","75"},
                { "151100310003","张飞","67","75","81"},
                { "151100310004","小乔","99","89","92"}
            };
            DataTable stuInfoTable = new DataTable
            {
                TableName = "学生信息表"
            };
            stuInfoTable.Columns.Add("学号", typeof(string));
            stuInfoTable.Columns.Add("姓名", typeof(string));
            stuInfoTable.Columns.Add("性别", typeof(string));
            stuInfoTable.Columns.Add("学院", typeof(string));
            stuInfoTable.Columns.Add("专业", typeof(string));
            stuInfoTable.Rows.Add("学号", "姓名", "性别", "学院", "专业");
            for (int i = 0; i < infos.GetLength(0); i++)
            {
                DataRow row = stuInfoTable.NewRow();
                for (int j = 0; j < infos.GetLength(1); j++)
                {
                    row[j] = infos[i, j];
                }
                stuInfoTable.Rows.Add(row);
            }
            ds.Tables.Add(stuInfoTable);

            DataTable stuScoreTable = new DataTable
            {
                TableName = "学生成绩表"
            };
            stuScoreTable.Columns.Add("学号", typeof(string));
            stuScoreTable.Columns.Add("姓名", typeof(string));
            stuScoreTable.Columns.Add("语文", typeof(string));
            stuScoreTable.Columns.Add("数学", typeof(string));
            stuScoreTable.Columns.Add("英语", typeof(string));
            stuScoreTable.Rows.Add("学号", "姓名", "语文", "数学", "英语");
            for (int i = 0; i < scores.GetLength(0); i++)
            {
                DataRow row = stuScoreTable.NewRow();
                for (int j = 0; j < scores.GetLength(1); j++)
                {
                    row[j] = scores[i, j];
                }
                stuScoreTable.Rows.Add(row);
            }
            ds.Tables.Add(stuScoreTable);
            return ds;
        }
    }
}
