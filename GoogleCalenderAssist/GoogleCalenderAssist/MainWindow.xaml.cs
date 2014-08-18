using System;
using System.IO;
using System.Data;
using System.Xml;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Google.GData.Client;
using Google.GData.Extensions;
using Google.GData.Calendar;
using Google.GData.AccessControl;

namespace GoogleCalenderAssist
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            //ファイルがあればインポートする
            if (File.Exists(System.Environment.CurrentDirectory + "/data.xml"))
            {
                XML xl = new XML();
                DataSet ds = new DataSet();
                ds = xl.Open(System.Environment.CurrentDirectory + "/data.xml");
                tbCal1.Text = ds.Tables[0].Rows[0][0].ToString();
                tbCal2.Text = ds.Tables[0].Rows[0][1].ToString();
                tbID.Text = ds.Tables[0].Rows[0][2].ToString();
                tbPASS.Text = ds.Tables[0].Rows[0][3].ToString(); 
            }
        }

        /// <summary>
        /// カレンダー作成
        /// </summary>
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            int year = 2014;

            for (int i = year; i < year + 7; i++)
            {
                EXCEL EX = new EXCEL();
                EX.Make();
                EX.EditSheet(i);
                EX.Save(i.ToString());
                EX.Close();
            }
        }

        /// <summary>
        /// 同期
        /// </summary>
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            //各設定保存
            XML xl = new XML();
            DataSet ds = new DataSet();
            DataTable dtb;

            dtb = ds.Tables.Add("data");
            dtb.Columns.Add("cal1",Type.GetType("System.String"));
            dtb.Columns.Add("cal2", Type.GetType("System.String"));
            dtb.Columns.Add("id", Type.GetType("System.String"));
            dtb.Columns.Add("pass", Type.GetType("System.String"));
            dtb.Rows.Add(new object[] {tbCal1.Text,tbCal2.Text,tbID.Text,tbPASS.Text});
            xl.Save(System.Environment.CurrentDirectory + "/data.xml", ds);

            // カレンダーサービスを作成
            CalendarService service = new CalendarService("companyName-applicationName-1");

            // 認証設定
            service.setUserCredentials(tbID.Text, tbPASS.Text);

            // 認証結果の確認
            try
            {
                // ここで例外がthrowされなければOK
                // 処理に少し時間がかかる
                var token = service.QueryClientLoginToken();
            }
            catch (InvalidCredentialsException ex)
            {
                // 認証に失敗している
                MessageBox.Show(ex.Message);
                service = null;
            }

            // 取得条件設定
            for (int year = 2014; year < 2021; year++)
            {
                //ファイルを開く
                EXCEL EX = new EXCEL();
                EX.Open(year.ToString());

                for (int calNo = 1; calNo < 3; calNo++)
                {
                    EventQuery query = new EventQuery();
                    if (calNo == 1)
                    {
                        query.Uri = new Uri("https://www.google.com/calendar/feeds/" + tbCal1.Text + "/private/full");
                    }
                    else
                    {
                        query.Uri = new Uri("https://www.google.com/calendar/feeds/" + tbCal2.Text + "/private/full");
                    }

                    query.StartTime = new DateTime(year, 1, 1);
                    query.EndTime = new DateTime(year, 12, 31);
                    query.SortOrder = CalendarSortOrder.descending;
                    //query.SingleEvents = true;


                    // 取得して表示
                    EventFeed feeds = service.Query(query);
                    IEnumerable<EventEntry> entries = feeds.Entries.Cast<EventEntry>();
                    foreach (EventEntry entry in entries)
                    {
                        //ファイルに書き込み
                        EX.Write(entry.Times.First().StartTime.Month, entry.Times.First().StartTime.Day,
                            entry.Times.First().StartTime.TimeOfDay.ToString(), entry.Locations.First().ValueString, entry.Title.Text, calNo);
                    }
                }

                //ファイル保存
                EX.Save(year.ToString());
                //ファイルを閉じる
                EX.Close();
            }
        }
    }
}
