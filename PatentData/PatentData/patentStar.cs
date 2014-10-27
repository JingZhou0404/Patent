using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Threading;
using mshtml;
namespace PatentData
{
    public partial class patentStar : Form
    {
        bool loading = true;
        string patentid = "";
        System.Timers.Timer mytimer = null;
        public patentStar()
        {
            InitializeComponent();

        }
        //爬取入口
        private void btnSearch_Click(object sender, EventArgs e)
        {
            string sql = "select  applyid from temp_flzt where applyid>'8183023'and patentid is  null  order by applyid ";
            SqlDataReader reader = SqlHelper.ExecuteReader(SqlHelper.ConnString, CommandType.Text, sql);
            while (reader.Read())
            {
                string applyid = reader["applyid"].ToString();
                searchPatentStar(applyid);
            }

        }


        //爬取数据
        public void searchPatentStar(string applyid)
        {
            try
            {
                string url = "http://www.patentstar.cn/cprs2010/";
                loading = true;
                webBrowser1.Navigate(url, false);
                while (loading)
                {
                    Application.DoEvents();
                }
                //找到页面中输入关键字的input标签
                HtmlElement keys = webBrowser1.Document.All["searchContent"];
                //为input标签输入value值
                keys.SetAttribute("value", applyid);

                //找到页面中button或者submit按钮
                HtmlElement hit = webBrowser1.Document.All["btnSearch"];
                //触发搜索按钮的click事件
                hit.InvokeMember("click");
                patentid = applyid;

                //启动时间控件
                mytimer = new System.Timers.Timer(500);
                mytimer.AutoReset = false;
                mytimer.Enabled = true;
                mytimer.Elapsed += new System.Timers.ElapsedEventHandler(Timer_Elapsed);
            }
            catch { }
        }
        public delegate void GetDataHandler();
        private void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            Invoke(new GetDataHandler(getData));
            mytimer.Start();
        }
        private void getData()
        {
            try
            {
                HtmlElement ele = webBrowser1.Document.GetElementById("cprsResTable");//用于标识数据是否已经加载出来的Dom元素
                if (ele == null)
                {
                    mytimer.Start();
                }
                else
                {
                    mytimer.Stop();
                    int count = ele.Children.Count;
                    //若是通过一个专利号查出多个记录直接忽略此专利号，仅查询只有一个记录的专利
                    if (count == 1)
                    {
                        HtmlElement elem = ele.Children[0];
                        string text = elem.InnerText;
                        text = text.Substring(text.IndexOf("申请号:") + 4, text.IndexOf("申请日") - text.IndexOf("申请号:") - 4).Trim();
                        patentid = text.Substring(0, text.Length - 2);
                        SqlHelper.ExecuteNonQuery(SqlHelper.ConnString, CommandType.Text, "update temp_flzt set patentid='" + text + "' where applyid='" + patentid + "'");
                    }
                }
            }
            catch { }

        }
        private void Timer_TimesUp(object sender, System.Timers.ElapsedEventArgs e)
        {
            //到达指定时间若还是没有加载出来网页框架，重新启动Timer，再重新等待网页加载
            HtmlElement flag = webBrowser1.Document.All["showModule"];
            if (flag == null)
                mytimer.Start();
        }
        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (webBrowser1.ReadyState == WebBrowserReadyState.Complete || (e.Url != webBrowser1.Url))//判断加载完毕
            {
                //自动点击弹出确认或弹出提示
                IHTMLDocument2 vDocument = (IHTMLDocument2)webBrowser1.Document.DomDocument;
                vDocument.parentWindow.execScript("function confirm(str){return true;} ", "javascript"); //弹出确认
                vDocument.parentWindow.execScript("function alert(str){return true;} ", "javaScript");//弹出提示
                loading = false;
                return;
            }
        }

        private void patentStar_Load(object sender, EventArgs e)
        {

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
