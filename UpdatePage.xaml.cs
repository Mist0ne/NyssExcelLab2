using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using System.Windows.Shapes;

namespace NyssExcelLab2
{
    /// <summary>
    /// Логика взаимодействия для UpdatePage.xaml
    /// </summary>
    public partial class UpdatePage : Window
    {
        public UpdatePage(ObservableCollection<ThreatModel> changedDataOld, ObservableCollection<ThreatModel> changedDataNew)
        {
            // Displaying data with differences in the window constructor.

            InitializeComponent();
            string oldStr = "";
            string changeStr = "";
            string newStr = "";
            for (int i = 0; i < changedDataOld.Count; i++)
            {
                changeStr += $"Change to the {i+1} entry:\n";

                if (changedDataOld[i].Id != changedDataNew[i].Id) { changeStr += $"Old Id: {changedDataOld[i].Id}\nNew Id: {changedDataNew[i].Id}\n"; }
                oldStr += $"{changedDataOld[i].Id}\n";
                newStr += $"{changedDataNew[i].Id}\n";

                if (changedDataOld[i].Name != changedDataNew[i].Name) { changeStr += $"Old Name: {changedDataOld[i].Name}\nNew Name: {changedDataNew[i].Name}\n"; }
                oldStr += $"{changedDataOld[i].Name}\n";
                newStr += $"{changedDataNew[i].Name}\n";

                if (changedDataOld[i].Capture != changedDataNew[i].Capture) { changeStr += $"Old Capture: {changedDataOld[i].Capture}\nNew Capture: {changedDataNew[i].Capture}\n"; }
                oldStr += $"{changedDataOld[i].Capture}\n";
                newStr += $"{changedDataNew[i].Capture}\n";

                if (changedDataOld[i].ThreatSource != changedDataNew[i].ThreatSource) { changeStr += $"Old ThreatSource: {changedDataOld[i].ThreatSource}\nNew ThreatSource: {changedDataNew[i].ThreatSource}\n"; }
                oldStr += $"{changedDataOld[i].ThreatSource}\n";
                newStr += $"{changedDataNew[i].ThreatSource}\n";

                if (changedDataOld[i].ThreatTarget != changedDataNew[i].ThreatTarget) { changeStr += $"Old ThreatTarget: {changedDataOld[i].ThreatTarget}\nNew ThreatTarget: {changedDataNew[i].ThreatTarget}\n"; }
                oldStr += $"{changedDataOld[i].ThreatTarget}\n";
                newStr += $"{changedDataNew[i].ThreatTarget}\n";

                if (changedDataOld[i].Confidentiality != changedDataNew[i].Confidentiality) { changeStr += $"Old Confidentiality: {changedDataOld[i].Confidentiality}\nNew Confidentiality: {changedDataNew[i].Confidentiality}\n"; }
                oldStr += $"{changedDataOld[i].Confidentiality}\n";
                newStr += $"{changedDataNew[i].Confidentiality}\n";

                if (changedDataOld[i].Integrity != changedDataNew[i].Integrity) { changeStr += $"Old Integrity: {changedDataOld[i].Integrity}\nNew Integrity: {changedDataNew[i].Integrity}\n"; }
                oldStr += $"{changedDataOld[i].Integrity}\n";
                newStr += $"{changedDataNew[i].Integrity}\n";

                if (changedDataOld[i].Availability != changedDataNew[i].Availability) { changeStr += $"Old Availability: {changedDataOld[i].Availability}\nNew Availability: {changedDataNew[i].Availability}\n"; }
                oldStr += $"{changedDataOld[i].Availability}\n";
                newStr += $"{changedDataNew[i].Availability}\n";

                changeStr += "\n";
                oldStr += "\n";
                newStr += "\n";
            }
            OldInfoText.Text = oldStr;
            NewInfoText.Text = newStr;
            ChangeListText.Text = changeStr;
        }
    }
}
