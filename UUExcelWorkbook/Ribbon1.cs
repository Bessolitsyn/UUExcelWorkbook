using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System.Runtime.Serialization.Json;
using System.Runtime.Serialization;
using System.IO;
using Newtonsoft.Json;

namespace UUExcelWorkbook
{
    public partial class Ribbon1
    {
        Application Application;
        

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            Application = Globals.ThisWorkbook.Application;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog.Filter = "JSON files (*.json)|*.json";
            saveFileDialog.FileOk += SaveFileDialog_FileOk;

            saveFileDialog.ShowDialog();
           


        }

        private void SaveFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Excel.Workbook workbook = Application.ActiveWorkbook;
            List<KeyValuePair<int, string>> FlowMeterPropNames = new List<KeyValuePair<int, string>>();
            List<FullFLowMeter> fullFLowMeters = new List<FullFLowMeter>();

            var range = Application.Range["title"];
            if (range != null)
            {
                var value = range[1, 1].Value;
                //System.Windows.Forms.MessageBox.Show(value);
                for (int i = 11; i < 22; i++)
                {
                    FlowMeterPropNames.Add(new KeyValuePair<int, string>(i, range[1, i].value));
                }

            }
            int key = 0;
            foreach (Excel.Name name in workbook.Names)
            {
                //System.Windows.Forms.MessageBox.Show(name.Name);

                if (name.Name.Contains("Data"))
                {
                    range = Application.Range[name.Name];
                    if (range != null)
                    {
                       
                        foreach (Range row in range.Rows)
                        {
                            var item = row.Cells;
                            //System.Windows.Forms.MessageBox.Show(item[1][1].Text);
                            var flowMeter = new FlowMeter();
                            flowMeter.Id = ++key;
                            flowMeter.Code = item[1, 1].Text;
                            flowMeter.Type = item[1, 2].Text;
                            flowMeter.ConnectionType = item[1, 3].Text;
                            flowMeter.Reducer = item[1, 4].Text;
                            flowMeter.DIA1 = item[1, 5].Text;
                            flowMeter.DIA0 = item[1, 6].Text;
                            flowMeter.CostPrice = item[1, 7].Text;
                            flowMeter.Price = item[1, 8].Text;
                            flowMeter.FlowRateMin = item[1, 9].Text;
                            flowMeter.FlowRateMax = item[1, 10].Text;
                            flowMeter.Name = item[1, 34].Text;
                            var flowMeterProperties = new List<FlowMeterProperty>();
                            //System.Windows.Forms.MessageBox.Show(value);
                            for (int i = 11; i < 22; i++)
                            {
                                var flowMeterProperty = new FlowMeterProperty();
                                flowMeterProperty.Value = item[1, i].Text;
                                flowMeterProperty.FlowMeterId= flowMeter.Id;
                                flowMeterProperty.Column = i.ToString();
                                flowMeterProperty.Name = FlowMeterPropNames.Single(vp => vp.Key == i).Value;
                                flowMeterProperties.Add(flowMeterProperty);
                            }
                            fullFLowMeters.Add(new FullFLowMeter() { FlowMeter = flowMeter, FlowMeterProperties = flowMeterProperties });

                        }


                        #region Header
                        // Наименование столбцов 
                        //-2
                        //3   C   Кодовый номер
                        //4   D   Тип РС
                        //5   E   Тип подключения
                        //6   F   Переход
                        //7   G   ДУ1
                        //8   H   ДУ0
                        //9   I   Себестоимость
                        //10  J   Прайсовая стоимость
                        //11  K   Qmin
                        //12  L   Qmax
                        //13  M   NOa1, °	
                        //14  N   a2P, °	
                        //15  O   ДУ2
                        //16  P   L1
                        //17  Q   L21
                        //18  R   L22
                        //19  S   L3
                        //20  T   Lрасх
                        //21  U   Lпр.уч*
                        //22  V   Число Рейнольдса Reсуж
                        //23  W   Коэффициент трения сужения
                        //24  X   Коэффициент сопротивления сужения
                        //25  Y   Число Рейнольдса Reпрям1
                        //26  Z   Коэффициент трения прямой участок
                        //27  AA  Число Рейнольдса Reрасш
                        //28  AB  Коэффициент трения расширения
                        //29  AC  Коэффициент сопротивления расширения
                        //30  AD  Значение потерь на прямом участке
                        //31  AE  Значение потерь на сужении
                        //32  AF  Значение потерь на расширении
                        //33  AG  Потери в расходомере(L-каналах Питерфлоу),  м.в.ст
                        //34  AH  Суммарные потери
                        //35  AI  Скорость потока в расходомере, м / с
                        #endregion

                    }

                }



            }

            DataContractJsonSerializer jsonFormatter = new DataContractJsonSerializer(typeof(List<FullFLowMeter>));

            string collection = JsonConvert.SerializeObject(fullFLowMeters);
            byte[] file = Encoding.UTF8.GetBytes(collection.ToCharArray());
            var fname = ((System.Windows.Forms.SaveFileDialog)sender).FileName;
            using (FileStream fs = new FileStream(fname, FileMode.OpenOrCreate))
            {
                fs.Write(file, 0, file.Length);
                fs.Close();
                //jsonFormatter.WriteObject(fs, fullFLowMeters);
            }
            //throw new NotImplementedException();
        }
    }
}
