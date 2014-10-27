using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;


namespace PatentData
{
    public partial class Excel : Form
    {
        public Excel()
        {
            InitializeComponent();
        }

        private void btnAddress_Click(object sender, EventArgs e)
        {
            string fileName = "";
            this.openFileDialog1.Filter = "Excel文档(*.xlsx)|*.xlsx|Excel文档(*.xls)|*.xls";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
                try
                {
                    LoadDataFromExcel(fileName);
                    MessageBox.Show("数据导入成功!", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("数据导入失败！失败原因:" + ex.Message, "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
        }

        //加载Excel
        public void LoadDataFromExcel(string filePath)
        {
            try
            {
                string strconn = "";
                int length = 0;
                strconn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1'";
                OleDbConnection OleConn = new OleDbConnection(strconn);
                OleConn.Open();
                string[] province = new string[] { "北京", "上海", "天津", "重庆", "河北", "山西", "河南", "辽宁", "吉林", "黑龙江", "内蒙古", "江苏", "山东", "安徽", "浙江", "福建", "湖北", "湖南", "广东", "广西", "江西", "四川", "海南", "贵州", "云南", "西藏", "陕西", "甘肃", "青海", "宁夏", "新疆" };
                string[] code = new string[] { "110000", "310000", "120000", "850000", "130000", "140000", "410000", "210000", "220000", "230000", "150000", "320000", "370000", "340000", "330000", "350000", "420000", "430000", "440000", "450000", "360000", "510000", "660000", "520000", "530000", "540000", "610000", "620000", "630000", "640000", "650000" };
                length = province.Length;
               
                for (int i = 0; i < length; i++)
                {
                    string sql = "select * from [" + province[i] + "$]";
                    OleDbDataAdapter OleDaExcel = new OleDbDataAdapter(sql, OleConn);
                    DataSet OleDsExcel = new DataSet();
                    OleDaExcel.Fill(OleDsExcel);
                    ExcelToDataBase(OleDsExcel, code[i], province[i]);
                }
                OleConn.Close();
            }
            catch(Exception e)
            {
                MessageBox.Show("数据绑定失败！失败原因:" + e.Message, "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        //处理Excel中的数据
        public void ExcelToDataBase(DataSet data,string code,string name)
        {
            try
            {

                string sql = "insert into t_zone(id,name,level,fatherid) values('" + code + "','" + name + "','1','0')";
                SqlHelper.ExecuteNonQuery(SqlHelper.ConnString, CommandType.Text, sql);
                DataTable dt = data.Tables[0];
                int count = dt.Rows.Count;
                int col = dt.Columns.Count;
                int code_city = 0;
                int code_area = 0;
                int j = 1;
                int k = 0;
                string sign = "";
                for (int i = 0; i < count; i++)
                {
                    if (col == 2)
                    {
                        code_city = int.Parse(code) + i + 100;
                        string sql_city = "insert into t_zone(id,name,level,fatherid) values('" + code_city + "','" + dt.Rows[i][1] + "','2','"+code+"')";
                        SqlHelper.ExecuteNonQuery(SqlHelper.ConnString, CommandType.Text, sql_city);
                    }
                    if (col == 3)
                    {
                        if (sign != dt.Rows[i][1].ToString())
                        {
                            k = 1;
                            code_city = int.Parse(code) + j*100;
                            code_area = code_city + k;
                            string sql_city = "insert into t_zone(id,name,level,fatherid) values('" + code_city + "','" + dt.Rows[i][1] + "','2','"+code+"')";
                            string sql_area = "insert into t_zone(id,name,level,fatherid) values('" + code_area + "','" + dt.Rows[i][2] + "','3','" + code_city + "')";
                            SqlHelper.ExecuteNonQuery(SqlHelper.ConnString, CommandType.Text, sql_city);
                            SqlHelper.ExecuteNonQuery(SqlHelper.ConnString, CommandType.Text, sql_area);
                            j++;
                            sign = dt.Rows[i][1].ToString();
                        }
                        else
                        {
                            k++;
                            code_area = code_city + k;
                            string sql_area = "insert into t_zone(id,name,level,fatherid) values('" + code_area + "','" + dt.Rows[i][2] + "','3','" + code_city + "')";
                            SqlHelper.ExecuteNonQuery(SqlHelper.ConnString, CommandType.Text, sql_area);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Excel数据导入数据库失败！失败原因:" + e.Message, "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnAgent_Click(object sender, EventArgs e)
        {
            string fileName = "";
            this.openFileDialog1.Filter = "Excel文档(*.xlsx)|*.xlsx|Excel文档(*.xls)|*.xls";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
                try
                {
                    LoadDataFromExcel_Agent(fileName);
                    MessageBox.Show("数据导入成功!", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("数据导入失败！失败原因:" + ex.Message, "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

            }
        }
        //加载Excel
        public void LoadDataFromExcel_Agent(string filePath)
        {
            try
            {
                string strconn = "";
                int length = 0;
                strconn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1'";
                OleDbConnection OleConn = new OleDbConnection(strconn);
                OleConn.Open();
                string[] province = new string[] { "北京", "上海", "天津", "重庆", "河北", "山西", "河南", "辽宁", "吉林", "黑龙江", "内蒙古", "江苏", "山东", "安徽", "浙江", "福建", "湖北", "湖南", "广东", "广西", "江西", "四川", "海南", "贵州", "云南", "陕西", "甘肃", "青海", "宁夏", "新疆", "香港", "其他专利代理机构" };
                //string[] code = new string[] { "110000", "310000", "120000", "850000", "130000", "140000", "410000", "210000", "220000", "230000", "150000", "320000", "370000", "340000", "330000", "350000", "420000", "430000", "440000", "450000", "360000", "510000", "660000", "520000", "530000", "540000", "610000", "620000", "630000", "640000", "650000" };
                length = province.Length;

                for (int i = 0; i < length; i++)
                {
                    string sql = "select * from [" + province[i] + "$]";
                    OleDbDataAdapter OleDaExcel = new OleDbDataAdapter(sql, OleConn);
                    DataSet OleDsExcel = new DataSet();
                    OleDaExcel.Fill(OleDsExcel);
                    ExcelToDataBaseAgent(OleDsExcel, province[i]);
                }
                OleConn.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show("数据绑定失败！失败原因:" + e.Message, "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }
        //代理机构
        public void ExcelToDataBaseAgent(DataSet data,string zone)
        {         
            string fatherid = "";
            string id = "";
            string name = "";
            string state = "";
            string address = "";
            string director = "";
            string contacts = "";
            string copartner = "";
            string agents = "";
            bool flag = false;
            bool flag_branch=false;
            DataTable dt = data.Tables[0];
            int count = dt.Rows.Count;
            for (int i = 0; i < count; i++)
            {
                string title=dt.Rows[i][0].ToString();
                switch(title)
                {
                    case "机 构 代 码":
                        fatherid = dt.Rows[i][1].ToString();
                        id = dt.Rows[i][1].ToString();
                        break;
                    case "机 构 名 称":
                        name = dt.Rows[i][1].ToString();
                        break;
                    case "当前法律状态":
                        state = dt.Rows[i][1].ToString();
                        break;
                    case "邮编 、地址":
                        address=dt.Rows[i][1].ToString();
                        break;
                    case "负　 责　人":
                        director = dt.Rows[i][1].ToString();
                        break;
                    case "联 系 电 话":
                        contacts = dt.Rows[i][1].ToString();
                        break;
                    case "合伙人/股东":
                        copartner = dt.Rows[i][1].ToString();
                        break;
                    case "代 　理　人":
                        agents = dt.Rows[i][1].ToString();
                        flag = true;
                        break;
                    case "分支机构代码":
                        id = dt.Rows[i][1].ToString();
                        break;
                    case"分支机构名称":
                        name = dt.Rows[i][1].ToString();
                        break;
                    case "分支机构地址":
                        address=dt.Rows[i][1].ToString();
                        break;
                    case "负 责 人":
                        director=dt.Rows[i][1].ToString();
                        break;
                    case "电 话":
                        contacts=dt.Rows[i][1].ToString();
                        break;
                    case "代 理 人":
                        director=dt.Rows[i][1].ToString();
                        flag_branch = true;
                        break;
                }
                try
                {
                    if (flag)
                    {
                        string sql = "insert into t_agent(id,name,state,address,director,contacts,copartner,agents,zone) values('" +
                            id + "','" + name + "','" + state + "','" + address + "','" + director + "','" + contacts + "','" + copartner + "','" + agents + "','" + zone + "')";
                        SqlHelper.ExecuteNonQuery(SqlHelper.ConnString, CommandType.Text, sql);
                        flag = false;
                        
                    }
                    if(flag_branch)
                    {
                        string sql = "insert into t_agent(fatherid,id,name,state,address,director,contacts,copartner,agents,zone) values('" +fatherid+"','"+
                           id + "','" + name + "','" + state + "','" + address + "','" + director + "','" + contacts + "','" + copartner + "','" + agents + "','" + zone + "')";
                        SqlHelper.ExecuteNonQuery(SqlHelper.ConnString, CommandType.Text, sql);
                        flag_branch = false;
                    }
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Excel导入数据库失败，失败原因:"+ex.Message,"提示信息",MessageBoxButtons.OK,MessageBoxIcon.Error);
                    return;
                }
            }

        }
    }
}
