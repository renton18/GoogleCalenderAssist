﻿using System;
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
            EXCEL EX = new EXCEL();
            EX.Make();
            EX.EditSheet(2014);
            EX.Save("2014");
            EX.Close();

        }

        /// <summary>
        /// 同期
        /// </summary>
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {

        }
    }
}