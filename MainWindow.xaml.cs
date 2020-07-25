using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
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

namespace NyssExcelLab2
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ObservableCollection<ThreatModel> custdata;
        ObservableCollection<ThreatModel> currentData;
        ObservableCollection<ShortThreatModel> shortData;
        bool _short = false;
        int page = 0;
        public int Iter { set; get; }



        public MainWindow()
        {
            InitializeComponent();
            custdata = new ObservableCollection<ThreatModel>();
            currentData = new ObservableCollection<ThreatModel>();
            shortData = new ObservableCollection<ShortThreatModel>();
        }

        private void OpenFile(object sender, RoutedEventArgs e)
        {
            custdata = ThreatDataFromExcel.GetData();

            if (custdata.Count >= 20)
            {
                for (int i = 0; i < 20; i++)
                {
                    currentData.Add(custdata[i]);
                    shortData.Add(new ShortThreatModel(custdata[i].Id, custdata[i].Name));
                }
            }
            else
            {
                try
                {
                    MessageBox.Show("This file isn't exist. I will to download it from the Internet");
                    WebClient webClient = new WebClient();
                    webClient.DownloadFile("https://bdu.fstec.ru/files/documents/thrlist.xlsx", Excel.GetPath() + "thrlist.xlsx");

                    OpenFile(sender, e);
                }
                catch
                {
                    MessageBox.Show("Perhaps you have internet problems((\nTry again later");
                }
            }
            
            DataFromExcel.ItemsSource = currentData; 
        }

        private void Left_Button_Click(object sender, RoutedEventArgs e)
        {
            if (page != 0)
            {
                page--;
                PageNum.Text = $"Page: {page+1}";
                int fromObject = (page) * 20;
                if (currentData.Count != 0 && custdata.Count != 0)
                {
                    currentData.Clear();
                    shortData.Clear();
                    for (int i = fromObject; i < (page + 1) * 20; i++)
                    {
                        currentData.Add(custdata[i]);
                        shortData.Add(new ShortThreatModel(custdata[i].Id, custdata[i].Name));
                    }
                }
            }
        }

        private void Right_Button_Click(object sender, RoutedEventArgs e)
        {
            if (currentData.Count != 0 && custdata.Count != 0)
            {
                if (currentData[currentData.Count-1] != custdata[custdata.Count-1])
                {
                    page++;
                    PageNum.Text = $"Page: {page+1}";
                    int lastObject = (page + 1) * 20;
                    if (currentData.Count == 20)
                    {
                        currentData.Clear();
                        shortData.Clear();
                        if (lastObject >= custdata.Count)
                        {
                            for (int i = page * 20; i < custdata.Count; i++)
                            {
                                currentData.Add(custdata[i]);
                                shortData.Add(new ShortThreatModel(custdata[i].Id, custdata[i].Name));
                            }
                        }
                        else
                        {
                            for (int i = page * 20; i < lastObject; i++)
                            {
                                currentData.Add(custdata[i]);
                                shortData.Add(new ShortThreatModel(custdata[i].Id, custdata[i].Name));
                            }
                        }
                    
                    }
                }
            }
            
        }


        private void ShortAllInfoChanged(object sender, RoutedEventArgs e)
        {
            _short = !_short;
            if (_short)
            {
                DataFromExcel.ItemsSource = shortData;
            }
            else
            {
                DataFromExcel.ItemsSource = currentData;
            }
        }


        private void ChangeLocalData(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (_short)
            {
                Iter = shortData.IndexOf((ShortThreatModel)e.Row.Item);
            }
            else
            {
                Iter = currentData.IndexOf((ThreatModel)e.Row.Item);
            }
        }

        private void ChangeDataInGrid(object sender, EventArgs e)
        {
            if (_short) {
                custdata[page * 20 + Iter].Id = shortData[Iter].Id;
                custdata[page * 20 + Iter].Name = shortData[Iter].Name;
                currentData[Iter].Id = shortData[Iter].Id;
                currentData[Iter].Name = shortData[Iter].Name;
            }
            else {
                custdata[page * 20 + Iter] = currentData[Iter];
                shortData[Iter].Id = currentData[Iter].Id;
                shortData[Iter].Name = currentData[Iter].Name;
            }
        }

        private void SaveFileAs(object sender, RoutedEventArgs e)
        {
            if (PathToSave.Text.Contains('/') || PathToSave.Text.Contains('\\') || PathToSave.Text.Contains(".xl") || PathToSave.Text == "")
            {
                MessageBox.Show($"Please, write ONLY name of file (without .xlsx), without / or \\ \nI'll save your data automatically in {Excel.GetPath()} *name of your file*.xlsx", "Error");
            }
            else
            {
                string right_path = Excel.GetPath() + PathToSave.Text + ".xlsx";

                Excel excel = new Excel();
                try
                {
                    excel.NewFile();
                    excel.WriteRange(1, 1, custdata.Count, 8, custdata);
                    excel.SaveAs(right_path);
                    MessageBox.Show("You data was saved in\n" + right_path);
                    excel = null;
                }
                catch
                {
                    MessageBox.Show("Some error was found. Please, check all information and try again. The file is most likely open in another program");
                }   
            }
        }

        private void UpdateData(object sender, RoutedEventArgs e)
        {
            ObservableCollection<ThreatModel> newData = ThreatDataFromExcel.GetData();
            ObservableCollection<ThreatModel> changedDataOld = new ObservableCollection<ThreatModel>();
            ObservableCollection<ThreatModel> changedDataNew = new ObservableCollection<ThreatModel>();


            int min = Math.Min(custdata.Count, newData.Count);
            for (int i = 0; i < min; i++)
            {
                if (custdata[i].Id != newData[i].Id || custdata[i].Name != newData[i].Name || custdata[i].Capture != newData[i].Capture || custdata[i].ThreatSource != newData[i].ThreatSource || custdata[i].ThreatTarget != newData[i].ThreatTarget || custdata[i].Confidentiality != newData[i].Confidentiality || custdata[i].Integrity != newData[i].Integrity || custdata[i].Availability != newData[i].Availability)
                {
                    changedDataOld.Add(custdata[i]);
                    changedDataNew.Add(newData[i]);
                    custdata[i] = newData[i];
                }
            }
            if (newData.Count > min)
            {
                for (int i = min; i < newData.Count; i++)
                {
                    changedDataNew.Add(newData[i]);
                }
            }
            UpdatePage up = new UpdatePage(changedDataOld, changedDataNew);
            up.Show();

            if (newData.Count >= 20)
            {
                for (int i = 0; i < 20; i++)
                {
                    currentData.Add(newData[i]);
                    shortData.Add(new ShortThreatModel(newData[i].Id, newData[i].Name));
                }
            }
            else
            {
                try
                {
                    MessageBox.Show("Mistake. The update was unsuccessful.\nThis file isn't exist. I will to download it from the Internet");
                    WebClient webClient = new WebClient();
                    webClient.DownloadFile("https://bdu.fstec.ru/files/documents/thrlist.xlsx", Excel.GetPath() + "thrlist.xlsx");

                    UpdateData(sender, e);
                }
                catch
                {
                    MessageBox.Show("Perhaps you have internet problems((\nTry again later");
                }
            }

            DataFromExcel.ItemsSource = currentData;
        }
    }
}
