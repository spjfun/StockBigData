﻿using HtmlAgilityPack;
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

            //抓取多個股票資料
            listBox2.Items.Add("2330");
            listBox2.Items.Add("2303");
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
            MemoryStream ms = new MemoryStream();

            //抓取多個股票資料
            for (int j = 0; j < listBox2.Items.Count; j++)
            {
                //將網頁來源資料暫存到記憶體內
                ms = new MemoryStream(url.DownloadData("http://tw.stock.yahoo.com/q/q?s=" + listBox2.Items[j].ToString()));

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
                    listBox1.Items.Add(nodeHeader.InnerText + ":" + txt[i].Trim().Replace("加到投資組合", "") + "");

                    //將 "加到投資組合" 這個字串過濾掉,其他的存入DB
                    SQLCommand += ", '" + txt[i].Trim().Replace("加到投資組合", "") + "'";

                    i++;
                }
                SQLCommand += ")";
                // SQL Command 結束

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
            }

            url = null;
            ms = null;

            //清除資料
            //doc = null;
            //hdc = null;
            //url = null;
            //ms.Close();
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

        //載入股票代碼
        private void button3_Click(object sender, EventArgs e)
        {
            LoadStocksCode();
        }

        //資料庫讀取資料
        private void LoadStocksCode()
        {
            var serverName = "127.0.0.1";
            var uidName = "root";
            var pwdName = "1234";
            var databaseName = "BDS";

            // 連線字串
            string connStr = String.Format("server={0};uid={1};pwd={2};database={3}",
            serverName, uidName, pwdName, databaseName);

            try
            {
                conn = new MySqlConnection(connStr);

                // 開啟連線
                conn.Open();

                GetDatabases();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Error connecting to the server: " + ex.Message);
            }


        }

        //資料庫讀取股票代碼
        private void GetDatabases()
        {
            MySqlDataReader reader = null;

            // SQL Command
            MySqlCommand cmd = new MySqlCommand("SELECT * FROM STOCKSCODE", conn);

            try
            {
                // 執行SQL 
                reader = cmd.ExecuteReader();

                listBox2.Items.Clear();
                while (reader.Read())
                {
                    listBox2.Items.Add(reader.GetString(0));
                }

            }
            catch (MySqlException ex)
            {
                // 錯誤訊息丟出 
                MessageBox.Show("Failed to populate database list: " + ex.Message);
            }
            finally
            {
                // 如果連線還沒關閉, 關閉它 
                if (reader != null) reader.Close();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            GetData2();
        }


        //取得爬蟲指定抓取資料
        //清空truncate table stockscode
        public void GetData2()
        {
            //指定來源網頁
            WebClient url = new WebClient();
            MemoryStream ms = new MemoryStream();


            //將網頁來源資料暫存到記憶體內
            ms = new MemoryStream(url.DownloadData("http://94im.com/thread-14545-1-1.html"));

            // 使用UTF8編碼讀入 HTML 
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.Load(ms, Encoding.UTF8);

            // 裝載第一層查詢結果 
            HtmlAgilityPack.HtmlDocument hdc = new HtmlAgilityPack.HtmlDocument();

            hdc.LoadHtml(doc.DocumentNode.SelectSingleNode("//table[@class='t_table'][1]").InnerHtml);

            var trlength = doc.DocumentNode.SelectNodes("//table[@class='t_table'][1]/tr").Count;


            List<HtmlNode> trlist = doc.DocumentNode.SelectNodes("//table[@class='t_table'][1]").Elements("tr").ToList();

            var sqlcmd = "INSERT INTO stockscode (Code, Name) VALUES ";

            for (int indextr = 1; indextr < trlist.Count; indextr++)
            {
                List<HtmlNode> s = trlist[indextr].Elements("td").ToList();

                var code = s[0].InnerText.Trim();
                var name = s[1].InnerText.Trim();

                if (indextr == trlength - 1)
                {
                    Console.WriteLine("indextr == trlength - 1");
                    Console.WriteLine(" code：" + code);
                    sqlcmd += "(" + code;
                    sqlcmd += ",'" + name + "');";
                    break;
                }
                else {
                    sqlcmd += "(" + code;
                    sqlcmd += ",'" + name + "'),";
                }

                Console.WriteLine(" indextr：" + indextr);
            }

            runsqlcmd(sqlcmd);

            Console.WriteLine("runsqlcmd(sqlcmd)");

            url = null;
            ms = null;

            //清除資料
            //doc = null;
            //hdc = null;
            //url = null;
            //ms.Close();

        }

        public void runsqlcmd(string sqlcmd)
        {
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
                cmd = new MySqlCommand(sqlcmd, conn);
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
            finally
            {
                if (conn != null) conn.Close();
            }

        }


    }
}
