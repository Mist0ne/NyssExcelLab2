using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NyssExcelLab2
{
    public class ThreatModel
    {
        public string Id { set; get; }
        public string Name { set; get; }
        public string Capture { set; get; }
        public string ThreatSource { set; get; }
        public string ThreatTarget { set; get; }
        public string Confidentiality { set; get; }
        public string Integrity { set; get; }
        public string Availability { set; get; }


        public ThreatModel(string id, string name, string capture, string threatSource, string threatTarget, string confidentiality, string integrity, string availability)
        {
            this.Id = id;
            this.Name = name;
            this.Capture = capture;
            this.ThreatSource = threatSource;
            this.ThreatTarget = threatTarget;
            if (confidentiality == "1"){this.Confidentiality = "Да";}
            else{this.Confidentiality = "Нет";}
            if (integrity == "1") { this.Integrity = "Да"; }
            else { this.Integrity = "Нет"; }
            if (availability == "1") { this.Availability = "Да"; }
            else { this.Availability = "Нет"; }
        }
    }


    public class ShortThreatModel
    {
        public string Id { set; get; }
        public string Name { set; get; }


        public ShortThreatModel(string id, string name)
        {
            this.Id = id;
            this.Name = name;
        }
    }


    public static class ThreatDataFromExcel
    {
        public static ObservableCollection<ThreatModel> GetData()
        {
            // Trying to return data from Excel file. Using ranges
            ObservableCollection<ThreatModel> ResultData = new ObservableCollection<ThreatModel>();
            try
            {
                Excel excel = new Excel(Excel.GetPath() + "thrlist.xlsx", 1);

                int starti = 2;
                int step = 50;
                bool flag = true;
                while (flag)
                {
                    string[,] info = excel.ReadRange(starti, 0, starti + step, 8);
                    starti += step;
                    for (int i = 0; i < step; i++)
                    {
                        if (info[i,0] == "")
                        {
                            flag = false;
                            break;
                        }
                        ResultData.Add(new ThreatModel("УБИ." + new string('0', 3 - info[i, 0].Length) + info[i, 0], info[i, 1], info[i, 2], info[i, 3], info[i, 4], info[i, 5], info[i, 6], info[i, 7]));
                    }
                }
                return ResultData;
            }

            // In case of an error, it will return an empty(or incompletely filled) object
            catch
            {
                return ResultData;
            }
        }
    }
}
