using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Heater_Cal_Demo_P4.Data;

namespace Heater_Cal_Demo_P4
{
    public class CsvManager
    {
        private static readonly object csvLock = new object();

        public static bool CsvSave(string logData) //엑셀 저장 함수
        {
            lock (csvLock)
            {
                try
                {
                    var appSettings = Settings.Instance;

                    string model = appSettings.Model_No?.Replace(",", " ") ?? "Model";
                    string lotRaw = appSettings.Lot_No?.Replace(",", " ") ?? "1";
                    string lot = $"LOT{lotRaw}";

                    string Date = DateTime.Now.ToString("yyyyMMdd");
                    string logDirectory = Path.Combine(Application.StartupPath, "LogFile", model, lot, Date);
                    Directory.CreateDirectory(logDirectory);

                    string todayFileName = DateTime.Now.ToString("yyyyMMdd") + $"_{appSettings.LogFileStack}" + ".csv";
                    string csvFilePath = Path.Combine(logDirectory, todayFileName);

                    bool fileExists = File.Exists(csvFilePath);

                    using (StreamWriter writer = new StreamWriter(csvFilePath, true, Encoding.UTF8))
                    {
                        if (!fileExists)
                        {

                            writer.Write($",Data Log");
                            writer.WriteLine();
                        }

                        //데이터 쓰기
                        //큰따옴표로 깜싸면 csv파일에서 텍스트로 인식함 => 전체를 텍스트로 쳐서, 정렬방식 통일
                        string sanitized = logData?.Replace(",", " ") ?? "";
                        writer.Write($",{sanitized}");
                        writer.WriteLine();
                    }

                    return true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"CSV 저장 오류: {ex.Message}");
                    return false;
                }
            }

        }
    }

}

