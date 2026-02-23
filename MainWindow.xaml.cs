﻿using System.Windows;
using System.Windows.Input;
using VUKVWeightApp.ViewModels;

namespace VUKVWeightApp
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            // DataContext nastav dřív, aby se bindingy inicializovaly rovnou s výchozími hodnotami (IP, texty atd.)
            DataContext = new MainViewModel();
            InitializeComponent();
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                ToggleMaximize();
                return;
            }
            DragMove();
        }

        private void Minimize_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void Maximize_Click(object sender, RoutedEventArgs e)
        {
            ToggleMaximize();
        }

        private void ToggleMaximize()
        {
            WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
