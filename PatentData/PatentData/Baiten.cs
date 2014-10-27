using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.Sql;
using mshtml;
namespace PatentData
{
    public partial class Baiten : Form
    {
        bool loading = true;
        public Baiten()
        {
            InitializeComponent();
        }
        private void Baiten_Load(object sender, EventArgs e)
        {

        }
        public void Search(string applyid)
        {
            string lawState = "";
            bool flag_law = true;
            try
            {
                DateTime time0 = DateTime.Now;
                //richTextBox1.Text = "0：" + Environment.NewLine;
                string url = "http://so.baiten.cn/results?q=" + applyid;
                loading = true;
                webBrowser1.Navigate(url, false);
                while (loading)
                {
                    Application.DoEvents();
                }
                TimeSpan time1 = DateTime.Now - time0;
                richTextBox1.Text = richTextBox1.Text + "01:" + time1.Seconds.ToString() + Environment.NewLine;
                HtmlElementCollection eles = webBrowser1.Document.GetElementsByTagName("li");
                foreach (HtmlElement ele in eles)
                {
                    string eleClass = "";

                    eleClass = ele.GetAttribute("className");

                    if (eleClass.Equals("sm-c-r-color"))
                    {
                        HtmlElementCollection lawEles = ele.GetElementsByTagName("span");
                        foreach (HtmlElement lawele in lawEles)
                        {
                            if (lawele.GetAttribute("className").Contains("lawState"))
                            {
                                lawState = lawState + lawele.InnerText + " ";
                            }
                        }
                        flag_law = false;
                        TimeSpan time2 = DateTime.Now - time0;
                        richTextBox1.Text = richTextBox1.Text + "1:" + time2.Seconds.ToString() + Environment.NewLine;
                        InsertData(applyid, lawState);
                        TimeSpan time3 = DateTime.Now - time0;
                        richTextBox1.Text = richTextBox1.Text + "2:" + time3.Seconds.ToString() + Environment.NewLine;
                        break;
                    }

                }
                if (flag_law)
                {
                    TimeSpan time4 = DateTime.Now - time0;
                    richTextBox1.Text = richTextBox1.Text + "3:" + time4.Seconds.ToString() + Environment.NewLine;
                    updateData(applyid);
                    TimeSpan time5 = DateTime.Now - time0;
                    richTextBox1.Text = richTextBox1.Text + "4:" + time5.Seconds.ToString() + Environment.NewLine;

                }
            }
            catch
            {

            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            string sql = "select patentid from flzt_law where enable=1 and flag=0 and tag is null order by patentid";
            SqlDataReader reader = SqlHelper.ExecuteReader(SqlHelper.ConnString, CommandType.Text, sql);
            while (reader.Read())
            {
                string patentid = reader["patentid"].ToString();
                Search(patentid);
            }
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (webBrowser1.ReadyState == WebBrowserReadyState.Complete || (e.Url != webBrowser1.Url))//判断加载完毕
            {
                IHTMLDocument2 vDocument = (IHTMLDocument2)webBrowser1.Document.DomDocument;
                vDocument.parentWindow.execScript("function confirm(str){return true;} ", "javascript"); //弹出确认
                vDocument.parentWindow.execScript("function alert(str){return true;} ", "javaScript");//弹出提示
                loading = false;
                return;
            }
        }

        //数据加入数据库
        public void InsertData(string patentid, string lawState)
        {
            string sql = "update flzt_law set tag='" + lawState + "' where patentid='" + patentid + "'";
            SqlHelper.ExecuteNonQuery(SqlHelper.ConnString, CommandType.Text, sql);
        }
        public void updateData(string patentid)
        {
            string sql = "update flzt_law set flag='1' where patentid='" + patentid + "'";
            SqlHelper.ExecuteNonQuery(SqlHelper.ConnString, CommandType.Text, sql);
        }

        private void webBrowser1_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            //自动点击弹出确认或弹出提示
            IHTMLDocument2 vDocument = (IHTMLDocument2)webBrowser1.Document.DomDocument;
            vDocument.parentWindow.execScript("function confirm(str){return true;} ", "javascript"); //弹出确认
            vDocument.parentWindow.execScript("function alert(str){return true;} ", "javaScript");//弹出提示
        }

    }
}
