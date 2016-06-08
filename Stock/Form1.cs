using HtmlAgilityPack;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Stock
{
    public partial class Form1 : Form
    {
        private MySqlConnection conn;
        private MySqlCommand cmd;
        private int nextTime;
        private IAsyncResult asyncResult;
        private DateTime start;

        public Form1()
        {
            InitializeComponent();
        }


        private void button1_Click_1(object sender, EventArgs e)
        {
            GetData();
        }

        //每秒處理, 當遇到 13:30:00, 執行getdata()
        private void timer1_Tick(object sender, EventArgs e)
        {
            var strTime = DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00");

            switch (strTime)
            {
                case "142400":
                    GetData();
                    break;
            }
        }

        //取得爬蟲指定抓取資料
        public void GetData()
        {
            //指定來源網頁
            WebClient url = new WebClient();
            MemoryStream ms = new MemoryStream(url.DownloadData("http://tw.stock.yahoo.com/q/q?s=2317"));
            //以奇摩股市為例http://tw.stock.yahoo.com //2317 表示為股票代碼);

            // 使用預設編碼讀入 HTML 
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.Load(ms, Encoding.Default);

            // 裝載第一層查詢結果 
            HtmlAgilityPack.HtmlDocument hdc = new HtmlAgilityPack.HtmlDocument();

            //XPath 來解讀它 /html[1]/body[1]/center[1]/table[2]/tr[1]/td[1]/table[1] 
            hdc.LoadHtml(doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/center[1]/table[2]/tr[1]/td[1]/table[1]").InnerHtml);

            // 取得個股標頭 
            HtmlNodeCollection htnode = hdc.DocumentNode.SelectNodes("./tr[1]/th");
            // 取得個股數值 
            string[] txt = hdc.DocumentNode.SelectSingleNode("./tr[2]").InnerText.Trim().Split('\n');
            int i = 0;

            // SQL Command 建立 
            var SQLCommand = "INSERT INTO STOCKS (Date, No, Time, Price, Buy, Sell, Fluctuation, Number, ClosePrice, ";
            SQLCommand += "OpenPrice, High, Low, StocksData ) VALUE ('" + DateTime.Now.ToString("yyyy/MM/dd") + "'";

            
            foreach (HtmlNode nodeHeader in htnode)
            {
                // 輸出資料 到 listBox
                listBox1.Items.Add(txt[i].Trim());

                //將 "加到投資組合" 這個字串過濾掉,其他的存入DB
                SQLCommand += ", '" + txt[i].Trim().Replace("加到投資組合", "") + "'";

                i++;
            }
            SQLCommand += ")";
            // SQL Command 結束

            //---------------------------------------------------------------------------------------------------
            var serverName = "127.0.0.1";
            var uidName = "root";
            var pwdName = "1234";
            var databaseName = "BDS";
            string connStr = String.Format("server={0};uid={1};pwd={2};database={3}",
            serverName, uidName, pwdName, databaseName);
            conn = new MySqlConnection(connStr);
            try
            {
                conn.Open();

                cmd = new MySqlCommand(SQLCommand, conn);
                cmd.ExecuteNonQuery();


                asyncResult = cmd.BeginExecuteNonQuery();
                nextTime = 5;
                //timer1.Enabled = true;
                start = DateTime.Now;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception: " + ex.Message);
            }


            //清除資料
            doc = null;
            hdc = null;
            url = null;
            ms.Close();
        }

        //連線MariaDB
        private void button2_Click_1(object sender, EventArgs e)
        {
            //---------------------------------------------------------------------------------------------------
            // 伺服器名稱
            var serverName = "127.0.0.1";
            // 帳號
            var uidName = "root";
            // 密碼
            var pwdName = "1234";
            // 資料庫
            var databaseName = "BDS";
            // 連線字串
            string connStr = String.Format("server={0};uid={1};pwd={2};database={3}",
            serverName, uidName, pwdName, databaseName);

            conn = new MySqlConnection(connStr);
            //conn = new MySqlConnection("Server=localhost;Database=BDS;Uid=root;Pwd=1234;");

            try
            {
                // 開啟連線
                conn.Open();

                // SQL Command
                string sql = "SELECT * FROM STOCKS";
                cmd = new MySqlCommand(sql, conn);

                // 執行SQL 
                cmd.ExecuteNonQuery();

                asyncResult = cmd.BeginExecuteNonQuery();
                nextTime = 5;
                start = DateTime.Now;
            }
            catch (Exception ex)
            {
                // 錯誤訊息丟出 
                MessageBox.Show("Exception: " + ex.Message);
            }

            // 如果連線還沒關閉, 關閉它 
            if (conn != null)
                conn.Close();

        }
    }
}
