using System;
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
        }

        /// <summary>
        /// カレンダー作成
        /// </summary>
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            int year = DateTime.Now.Year;

            for (int i = year; i < year + 10; i++)
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

            // 取得条件設定（開始時間が2011年3月19日の予定を降順で取得）
            EventQuery query = new EventQuery();
            query.Uri = new Uri("https://www.google.com/calendar/feeds/default/private/full");

            int year = 2014;
            query.StartTime = new DateTime(year, 1, 1);
            query.EndTime = new DateTime(year, 12, 31);
            query.SortOrder = CalendarSortOrder.descending;
            //query.SingleEvents = true;

            //ファイルを開く
            EXCEL EX = new EXCEL();
            EX.Open(year.ToString());

            // 取得して表示
            EventFeed feeds = service.Query(query);
            IEnumerable<EventEntry> entries = feeds.Entries.Cast<EventEntry>();
            foreach (EventEntry entry in entries)
            {              
                //ファイルに書き込み
                EX.Write(entry.Times.First().StartTime.Month,entry.Times.First().StartTime.Day,
                    entry.Times.First().StartTime.TimeOfDay.ToString(),entry.Locations.First().ValueString,entry.Title.Text);
            }
            //ファイル保存
            EX.Save(year.ToString());
            //ファイルを閉じる
            EX.Close();
        }
    }
}
