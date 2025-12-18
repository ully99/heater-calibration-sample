using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO.Ports;
using System.Windows.Forms;
using Heater_Cal_Demo_P4.Communication;
using Heater_Cal_Demo_P4.Data;
using Modbus.Device;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ZedGraph;
using System.Diagnostics;





namespace Heater_Cal_Demo_P4
{

    public partial class MainForm : Form
    {
        #region Field
        private SerialChannelPort DevicePort = null;
        private MeasuringChannelPort MeasPort = null;
        int Device_ReadTimeOut_Value = 3000; //디바이스 수신 타임아웃
        int Meas_ReadTimeOut_value = 100; //온도계 수신 타임아웃

        //graph
        RollingPointPairList _listA = new RollingPointPairList(6000);
        RollingPointPairList _listP = new RollingPointPairList(6000);
        RollingPointPairList _listExt = new RollingPointPairList(6000);
        LineItem _curveA, _curveP, _curveExt;

        RollingPointPairList _listA_Veri = new RollingPointPairList(6000);
        RollingPointPairList _listP_Veri = new RollingPointPairList(6000);
        RollingPointPairList _listExt_Veri = new RollingPointPairList(6000);
        LineItem _curveA_Veri, _curveP_Veri, _curveExt_Veri;

        readonly Stopwatch _sw = new Stopwatch();
        const int X_WINDOW = 3000;  // x축 0~3000 (라벨: Time (x100ms))
        const int X_WINDOW_VERI = 3500; //Verification용
        //폴링 주기
        private readonly System.Windows.Forms.Timer _pollTimer90 = new System.Windows.Forms.Timer() { Interval = 90 };
        private readonly System.Windows.Forms.Timer _pollTimer90_Veri = new System.Windows.Forms.Timer() { Interval = 90 };

        //포트 동시 접근 보호
        private readonly SemaphoreSlim _ioGate = new SemaphoreSlim(1, 1);

        //최근 값 캐시
        private readonly object _latestLock = new object();
        private double? _lastTempP, _lastTempA, _lastMaurer;
        private double _lastRunTimeSec;
        private ushort _lastProfileStep;

        //실행 취소용
        private CancellationTokenSource _runCts;

        // == 스케줄: 특정 시각(경과 ms)에 디바이스로 Write 실행 ==
        private class ScheduledAction
        {
            public long DueMs;                               // 실행 예정 시각 (stopwatch 기준 ms)
            public Func<CancellationToken, Task> Action;    // 실행할 비동기 작업
            public bool Done = false;
        }
        private readonly List<ScheduledAction> _schedule = new List<ScheduledAction>();

        private static readonly (bool ok, double tp, double ta, double runSec, ushort step) DEV_EMPTY = (false, 0, 0, 0, 0);

        // === Point3 ===
        bool _p3Active = false;
        //bool _p3DiffWritten = false;
        bool _p3Averaged = false;

        private readonly List<double> _p3TempP = new List<double>();
        private readonly List<double> _p3Thermo = new List<double>();

        const double P3_SAMPLE_START = 29.0;
        const double P3_SAMPLE_END = 30.0;

        public float X3 { get; private set; } = float.NaN;
        public float Y3 { get; private set; } = float.NaN;

        // ===Point2 ===
        bool _p2Active = false;
        //bool _p2DiffWritten = false;
        bool _p2Averaged = false;

        private readonly List<double> _p2TempP = new List<double>();
        private readonly List<double> _p2Thermo = new List<double>();

        const double P2_SAMPLE_START = 59.0;
        const double P2_SAMPLE_END = 60.0;

        public float X2 { get; private set; } = float.NaN;
        public float Y2 { get; private set; } = float.NaN;

        // ===Point1 ===
        bool _p1Active = false;
        //bool _p1DiffWritten = false;
        bool _p1Averaged = false;

        private readonly List<double> _p1TempP = new List<double>();
        private readonly List<double> _p1Thermo = new List<double>();

        const double P1_SAMPLE_START = 89.0;
        const double P1_SAMPLE_END = 90.0;

        public float X1 { get; private set; } = float.NaN;
        public float Y1 { get; private set; } = float.NaN;

        //후속 작업
        const long ABC_Write_Time = 92_000;
        const long Set_Verification_Profile_Time = 94_000;
        const long Set_TR_Mode_Time = 95_000;
        const long Check_TR_Mode_Time = 96_000;
        const long Set_Calibration_Flag_TIme = 97_000;
        const long Check_Verification_Profile_Time = 98_000;
        const long Flash_Update_Time = 99_000;

        //결과 플래그
        public short Cal_Result_flag;

        //Verification 작업 
        //const double SECTION_1_START = 23.0;
        //const double SECTION_1_END = 24.0;
        private readonly List<double> _S1MaurerTemp = new List<double>();
        public float _S1Data { get; private set; } = float.NaN;
        bool _S1Active = false;
        bool _S1Averaged = false;

        //const double SECTION_2_START = 27.0;
        //const double SECTION_2_END = 28.0;
        private readonly List<double> _S2MaurerTemp = new List<double>();
        public float _S2Data { get; private set; } = float.NaN;
        bool _S2Active = false;
        bool _S2Averaged = false;

        //const double SECTION_3_START = 32.0;
        //const double SECTION_3_END = 33.0;
        private readonly List<double> _S3MaurerTemp = new List<double>();
        public float _S3Data { get; private set; } = float.NaN;
        bool _S3Active = false;
        bool _S3Averaged = false;

        //const double SECTION_4_START = 67.0;
        //const double SECTION_4_END = 68.0;
        private readonly List<double> _S4MaurerTemp = new List<double>();
        public float _S4Data { get; private set; } = float.NaN;
        bool _S4Active = false;
        bool _S4Averaged = false;

        //const double SECTION_5_START = 117.0;
        //const double SECTION_5_END = 118.0;
        private readonly List<double> _S5MaurerTemp = new List<double>();
        public float _S5Data { get; private set; } = float.NaN;
        bool _S5Active = false;
        bool _S5Averaged = false;

        //const double SECTION_6_START = 147.0;
        //const double SECTION_6_END = 148.0;
        private readonly List<double> _S6MaurerTemp = new List<double>();
        public float _S6Data { get; private set; } = float.NaN;
        bool _S6Active = false;
        bool _S6Averaged = false;

        //const double SECTION_7_START = 177.0;
        //const double SECTION_7_END = 178.0;
        private readonly List<double> _S7MaurerTemp = new List<double>();
        public float _S7Data { get; private set; } = float.NaN;
        bool _S7Active = false;
        bool _S7Averaged = false;

        //const double SECTION_8_START = 207.0;
        //const double SECTION_8_END = 208.0;
        private readonly List<double> _S8MaurerTemp = new List<double>();
        public float _S8Data { get; private set; } = float.NaN;
        bool _S8Active = false;
        bool _S8Averaged = false;

        //const double SECTION_9_START = 237.0;
        //const double SECTION_9_END = 238.0;
        private readonly List<double> _S9MaurerTemp = new List<double>();
        public float _S9Data { get; private set; } = float.NaN;
        bool _S9Active = false;
        bool _S9Averaged = false;

        //const double SECTION_10_START = 267.0;
        //const double SECTION_10_END = 268.0;
        private readonly List<double> _S10MaurerTemp = new List<double>();
        public float _S10Data { get; private set; } = float.NaN;
        bool _S10Active = false;
        bool _S10Averaged = false;

        //const double SECTION_11_START = 297.0;
        //const double SECTION_11_END = 298.0;
        private readonly List<double> _S11MaurerTemp = new List<double>();
        public float _S11Data { get; private set; } = float.NaN;
        bool _S11Active = false;
        bool _S11Averaged = false;

        //const double SECTION_12_START = 327.0;
        //const double SECTION_12_END = 328.0;
        private readonly List<double> _S12MaurerTemp = new List<double>();
        public float _S12Data { get; private set; } = float.NaN;
        bool _S12Active = false;
        bool _S12Averaged = false;

        //결과 플래그
        public short Veri_Result_flag;

        //후속 작업
        const long Set_TR_Mode_Time_Veri = 331_000;
        const long Set_Profile_Time_Veri = 332_000;
        const long Check_Profile_Time_Veri = 333_000;
        const long Set_Flag_Time_Veri = 334_000;
        const long Flash_Update_Time_Veri = 335_000;

        //DB
        private enum TaskKind { Calibration, Verification }
        private TaskKind _taskKind;

        private enum MesConnState { Off, Ok, Fail }

        //강제로 STOP할 경우 
        private bool deliberate_stop = false;
        
        

        #endregion

        #region Init
        public MainForm()
        {
            InitializeComponent();
            _pollTimer90.Tick += PollTimer90_Tick; //90ms 타이머 이벤트 연결
            _pollTimer90_Veri.Tick += PollTimer90_Veri_Tick;
        }

   

        private void MainForm_Load(object sender, EventArgs e)
        {
            RescanSerialPort(); //시리얼 포트 스켄
            InitTempGraph(); //Cal 그래프 세팅
            InitTempGraph_Veri(); //Veri 그래프 세팅
            DateStackUpOrInit(); //Log 파일 스텍 넘버
            Init_LogFileUse_RadioBtn(); //Log파일 사용 유무 라이오 버튼 초기화
            SettingMesImage(); //MES 이미지 파일 초기세팅
        }

        private void btnDbSettingOpen_Click(object sender, EventArgs e)
        {
            DBSettingForm form = new DBSettingForm(this);
            form.ShowDialog();
            
        }

        private void btnTestSettingOpen_Click(object sender, EventArgs e)
        {
            TestSettingForm form = new TestSettingForm(this);
            form.ShowDialog();
        }


        private void lblClearDeviceLog_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            tboxDeviceLog.Clear();
        }

        private void lblClearDeviceComm_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            tboxDeviceComm.Clear();
        }

        private void lblClearMeasLog_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            tboxMeasLog.Clear();
        }

        private void lblClearMeasComm_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            tboxMeasComm.Clear();
        }

        private void lblClearCalLog_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            tboxCalLog.Clear();
        }

        private void InitTempGraph()
        {
            var zg = TempGraph;
            var p = zg.GraphPane;

            // 타이틀/축
            p.Title.IsVisible = false;
            p.XAxis.Title.Text = "Time (x100ms)";
            p.YAxis.Title.Text = "Temp (°C)";

            // 배경
            p.Chart.Fill = new Fill(Color.White);
            p.Fill = new Fill(Color.White);

            // 범례
            p.Legend.IsVisible = true;
            p.Legend.Position = LegendPos.Top;      // 상단 중앙
            p.Legend.FontSpec.Size = 11;

            // 축 범위
            p.XAxis.Scale.Min = 0;
            p.XAxis.Scale.Max = X_WINDOW;           // 0~3000
            p.XAxis.Scale.MajorStep = 100;          // 눈금 간격
            p.XAxis.Scale.MinorStep = 20;

            p.YAxis.Scale.Min = 0;
            p.YAxis.Scale.Max = 350;                // 0~350°C
            p.YAxis.Scale.MajorStep = 50;
            p.YAxis.Scale.MinorStep = 10;

            // 그리드(빨간 점선 느낌)
            p.XAxis.MajorGrid.IsVisible = true;
            p.YAxis.MajorGrid.IsVisible = true;
            p.XAxis.MinorGrid.IsVisible = true;
            p.YAxis.MinorGrid.IsVisible = true;

            Color grid = Color.FromArgb(210, 70, 70);
            p.XAxis.MajorGrid.Color = grid;
            p.YAxis.MajorGrid.Color = grid;
            p.XAxis.MinorGrid.Color = Color.FromArgb(180, 220, 220, 220);
            p.YAxis.MinorGrid.Color = Color.FromArgb(180, 220, 220, 220);
            p.XAxis.MajorGrid.DashOn = 1; p.XAxis.MajorGrid.DashOff = 4;
            p.YAxis.MajorGrid.DashOn = 1; p.YAxis.MajorGrid.DashOff = 4;

            // 커브 추가 (Duty 제외)
            _curveA = p.AddCurve("TempA(°C)", _listA, Color.Red, SymbolType.None);
            _curveP = p.AddCurve("TempP(°C)", _listP, Color.Orange, SymbolType.None);
            _curveExt = p.AddCurve("Maurer Temp(°C)", _listExt, Color.MediumPurple, SymbolType.None);

            foreach (var c in new[] { _curveA, _curveP, _curveExt })
            {
                c.Line.Width = 2.0f;
                c.Line.IsSmooth = false;
                c.Symbol.IsVisible = false;
            }

            zg.IsAntiAlias = true;
            zg.AxisChange();
            zg.Invalidate();

            _sw.Restart();   // 시간 기준 시작
        }

        private void InitTempGraph_Veri()
        {
            var zg = TempGraph_Veri;
            var p = zg.GraphPane;

            // 타이틀/축
            p.Title.IsVisible = false;
            p.XAxis.Title.Text = "Time (x100ms)";
            p.YAxis.Title.Text = "Temp (°C)";

            // 배경
            p.Chart.Fill = new Fill(Color.White);
            p.Fill = new Fill(Color.White);

            // 범례
            p.Legend.IsVisible = true;
            p.Legend.Position = LegendPos.Top;      // 상단 중앙
            p.Legend.FontSpec.Size = 11;

            // 축 범위
            p.XAxis.Scale.Min = 0;
            p.XAxis.Scale.Max = X_WINDOW_VERI;           // 3500
            p.XAxis.Scale.MajorStep = 100;          // 눈금 간격
            p.XAxis.Scale.MinorStep = 20;

            p.YAxis.Scale.Min = 0;
            p.YAxis.Scale.Max = 350;                // 0~350°C
            p.YAxis.Scale.MajorStep = 50;
            p.YAxis.Scale.MinorStep = 10;

            // 그리드(빨간 점선 느낌)
            p.XAxis.MajorGrid.IsVisible = true;
            p.YAxis.MajorGrid.IsVisible = true;
            p.XAxis.MinorGrid.IsVisible = true;
            p.YAxis.MinorGrid.IsVisible = true;

            Color grid = Color.FromArgb(210, 70, 70);
            p.XAxis.MajorGrid.Color = grid;
            p.YAxis.MajorGrid.Color = grid;
            p.XAxis.MinorGrid.Color = Color.FromArgb(180, 220, 220, 220);
            p.YAxis.MinorGrid.Color = Color.FromArgb(180, 220, 220, 220);
            p.XAxis.MajorGrid.DashOn = 1; p.XAxis.MajorGrid.DashOff = 4;
            p.YAxis.MajorGrid.DashOn = 1; p.YAxis.MajorGrid.DashOff = 4;

            // 커브 추가 (Duty 제외)
            _curveA_Veri = p.AddCurve("TempA(°C)", _listA_Veri, Color.Red, SymbolType.None);
            _curveP_Veri = p.AddCurve("TempP(°C)", _listP_Veri, Color.Orange, SymbolType.None);
            _curveExt_Veri = p.AddCurve("Maurer Temp(°C)", _listExt_Veri, Color.MediumPurple, SymbolType.None);

            foreach (var c in new[] { _curveA_Veri, _curveP_Veri, _curveExt_Veri })
            {
                c.Line.Width = 2.0f;
                c.Line.IsSmooth = false;
                c.Symbol.IsVisible = false;
            }

            zg.IsAntiAlias = true;
            zg.AxisChange();
            zg.Invalidate();

            _sw.Restart();   // 시간 기준 시작
        }

        private void ResetPoint3()
        {
            _p3Active = true;
            _p3Averaged = false;
            _p3TempP.Clear();
            _p3Thermo.Clear();
            X3 = float.NaN;
            Y3 = float.NaN;
        }

        private void ResetPoint2()
        {
            _p2Active = true;
            _p2Averaged = false;
            _p2TempP.Clear();
            _p2Thermo.Clear();
            X2 = float.NaN;
            Y2 = float.NaN;
        }

        private void ResetPoint1()
        {
            _p1Active = true;
            _p1Averaged = false;
            _p1TempP.Clear();
            _p1Thermo.Clear();
            X1 = float.NaN;
            Y1 = float.NaN;
        }

        private void ResetSection1()
        {
            _S1Active = true;
            _S1Averaged = false;
            _S1MaurerTemp.Clear();
            _S1Data = float.NaN;
        }

        private void ResetSection2()
        {
            _S2Active = true;
            _S2Averaged = false;
            _S2MaurerTemp.Clear();
            _S2Data = float.NaN;
        }

        private void ResetSection3()
        {
            _S3Active = true;
            _S3Averaged = false;
            _S3MaurerTemp.Clear();
            _S3Data = float.NaN;
        }

        private void ResetSection4()
        {
            _S4Active = true;
            _S4Averaged = false;
            _S4MaurerTemp.Clear();
            _S4Data = float.NaN;
        }

        private void ResetSection5()
        {
            _S5Active = true;
            _S5Averaged = false;
            _S5MaurerTemp.Clear();
            _S5Data = float.NaN;
        }

        private void ResetSection6()
        {
            _S6Active = true;
            _S6Averaged = false;
            _S6MaurerTemp.Clear();
            _S6Data = float.NaN;
        }

        private void ResetSection7()
        {
            _S7Active = true;
            _S7Averaged = false;
            _S7MaurerTemp.Clear();
            _S7Data = float.NaN;
        }

        private void ResetSection8()
        {
            _S8Active = true;
            _S8Averaged = false;
            _S8MaurerTemp.Clear();
            _S8Data = float.NaN;
        }

        private void ResetSection9()
        {
            _S9Active = true;
            _S9Averaged = false;
            _S9MaurerTemp.Clear();
            _S9Data = float.NaN;
        }

        private void ResetSection10()
        {
            _S10Active = true;
            _S10Averaged = false;
            _S10MaurerTemp.Clear();
            _S10Data = float.NaN;
        }

        private void ResetSection11()
        {
            _S11Active = true;
            _S11Averaged = false;
            _S11MaurerTemp.Clear();
            _S11Data = float.NaN;
        }

        private void ResetSection12()
        {
            _S12Active = true;
            _S12Averaged = false;
            _S12MaurerTemp.Clear();
            _S12Data = float.NaN;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try { _pollTimer90?.Stop(); } catch { }
            try { if (_runCts != null) { _runCts.Cancel(); _runCts.Dispose(); _runCts = null; } } catch { }
            try { if (DevicePort?.IsOpen == true) DevicePort.Disconnect(); } catch { }
            try { if (MeasPort?.IsOpen == true) MeasPort.Disconnect(); } catch { }    
        }

        private void DateStackUpOrInit()
        {
            
            string CurrentDate = DateTime.Now.ToString("yyyyMMdd");

            Console.WriteLine($"CurrentDate :{CurrentDate} SavedDate : {Settings.Instance.SavedCurrentDate} Stack : {Settings.Instance.LogFileStack}");
            if (string.IsNullOrEmpty(Settings.Instance.SavedCurrentDate))
            {
                Settings.Instance.LogFileStack = 1;
                Settings.Instance.SavedCurrentDate = CurrentDate;
                Console.WriteLine($"저장된 날짜가 없으므로 현재날짜 할당 {Settings.Instance.SavedCurrentDate}");

                
                return;
            }

            if (CurrentDate == Settings.Instance.SavedCurrentDate)
            {
                Settings.Instance.LogFileStack++;
                Console.WriteLine($"날짜 동일 LogFileStack + 1 {Settings.Instance.LogFileStack}");
            }
            else
            {
                Settings.Instance.LogFileStack = 1;
                Settings.Instance.SavedCurrentDate = CurrentDate;
                Console.WriteLine($"날짜가 바뀌었으므로 LogFileStack 초기화 {Settings.Instance.LogFileStack}");
            }

            Settings.Instance.Save();
        }

        private void Init_LogFileUse_RadioBtn()
        {
            if(Settings.Instance.Use_Write_Log)
            {
                RbtnLogFileUseOn.Checked = true;
                RbtnLogFileUseOff.Checked = false;
            }
            else
            {
                RbtnLogFileUseOn.Checked = false;
                RbtnLogFileUseOff.Checked = true;
            }
        }

        private void RbtnLogFileUseOn_CheckedChanged(object sender, EventArgs e)
        {
            if (RbtnLogFileUseOn.Checked)
            {
                Settings.Instance.Use_Write_Log = true;
                Settings.Instance.Save();
            }
            else
            {
                Settings.Instance.Use_Write_Log = false;
                Settings.Instance.Save();
            }
                
        }

        private void SettingMesImage()
        {
            // 상태바 기본 설정
            statusStripMes.ImageScalingSize = new Size(16, 16); // ← System.Drawing 필요
            statusStripMes.ShowItemToolTips = true;

            CheckMesOnceAsync();
        }

        public async void CheckMesOnceAsync()
        {
            if (!Settings.Instance.USE_MES)
            {
                SetMesStatus(MesConnState.Off, "MES 비활성화");
                return;
            }

            SetMesStatus(MesConnState.Fail);

            bool ok = await Task.Run(() =>
            {
                try
                {
                    // 네 세팅값을 리소스 이용해 연결문자열 구성
                    string cs =
                        $@"Provider=SQLOLEDB;Data Source={Settings.Instance.DB_IP},{Settings.Instance.DB_PORT};" +
                        $@"Initial Catalog={Settings.Instance.DB_NAME};User ID={Settings.Instance.DB_USER};" +
                        $@"Password={Settings.Instance.DB_PW};Connect Timeout=5;";

                    using (var conn = new OleDbConnection(cs))
                    {
                        conn.Open(); // 붙기만 확인
                    }
                    return true;
                }
                catch
                {
                    return false;
                }
            });

            SetMesStatus(ok ? MesConnState.Ok : MesConnState.Fail);

            
        }

        private void SetMesStatus(MesConnState state, string tip = null)
        {
            switch (state)
            {
                case MesConnState.Off:
                    lblMesStatus.Image = Properties.Resources.LED_OFF_SM;
                    lblMesStatus.Text = "MES OFF";
                    break;

                case MesConnState.Ok:
                    lblMesStatus.Image = Properties.Resources.LED_GREEN_SM;
                    lblMesStatus.Text = "MES ON";
                    break;

                case MesConnState.Fail:
                    lblMesStatus.Image = Properties.Resources.LED_RED_SM;
                    lblMesStatus.Text = "MES Connect Fail";
                    break;
            }
            if (!string.IsNullOrEmpty(tip))
                lblMesStatus.ToolTipText = tip;
        }

        private void btnFtpSettingOpen_Click(object sender, EventArgs e)
        {
            FtpSettingForm form = new FtpSettingForm(this);
            form.ShowDialog();
        }




        public void Update_Setting_Field()
        {
            
        }


        #endregion

        #region Utility




        private void LogToUI(string Case, string msg, int status = 2) //Log textbox 함수
        {

            RichTextBox tbox;

            switch (Case)
            {
                case "Device":
                    tbox = tboxDeviceLog;
                    break;
                case "Maurer":
                    tbox = tboxMeasLog;
                    break;
                case "Cal":
                    tbox = tboxCalLog;
                    break;
                case "Veri":
                    tbox = tboxVeriLog;
                    break;
                default:
                    tbox = null;
                    break;
            }

            if (tbox == null) return;

            if (tbox.InvokeRequired)
            {
                tbox.Invoke(new Action(() =>
                {
                    if (status == 0) //실패
                    {
                        tbox.SelectionColor = Color.Red;
                        tbox.AppendText($"{Environment.NewLine} {msg} [{DateTime.Now:HH:mm:ss}]{Environment.NewLine}");
                        tbox.ScrollToCaret();
                    }
                    else if (status == 1) //성공
                    {
                        tbox.SelectionColor = Color.Green;
                        tbox.AppendText($"{Environment.NewLine} {msg} [{DateTime.Now:HH:mm:ss}]{Environment.NewLine}");
                        tbox.ScrollToCaret();
                    }
                    else
                    {
                        tbox.SelectionColor = Color.Black;
                        tbox.AppendText($"{Environment.NewLine} {msg} [{DateTime.Now:HH:mm:ss}]{Environment.NewLine}");
                        tbox.ScrollToCaret();
                    }

                }));
            }
            else
            {
                if (status == 0) //실패
                {
                    tbox.SelectionColor = Color.Red;
                    tbox.AppendText($"{Environment.NewLine} {msg} [{DateTime.Now:HH:mm:ss}]{Environment.NewLine}");
                    tbox.ScrollToCaret();
                }
                else if (status == 1) //성공
                {
                    tbox.SelectionColor = Color.Green;
                    tbox.AppendText($"{Environment.NewLine} {msg} [{DateTime.Now:HH:mm:ss}]{Environment.NewLine}");
                    tbox.ScrollToCaret();
                }
                else
                {
                    tbox.SelectionColor = Color.Black;
                    tbox.AppendText($"{Environment.NewLine} {msg} [{DateTime.Now:HH:mm:ss}]{Environment.NewLine}");
                    tbox.ScrollToCaret();
                }
            }

           
        }

        private void LogCommToUI(RichTextBox textbox, string msg, bool isTx)
        {
            var tbox = textbox;
            string prefix = isTx ? "[Tx]" : "[Rx]";
            string visible = msg.Replace("\r", "\\r").Replace("\n", "\\n");

            Action ui = () =>
            {
                tbox.SelectionColor = isTx ? Color.Blue : Color.Red;
                tbox.AppendText($"\r\n{prefix} : {visible} [{DateTime.Now:HH:mm:ss}]\r\n");
                tbox.ScrollToCaret();
            };
            if (tbox.InvokeRequired) tbox.Invoke(ui); else ui();
        }

        private void LogOleDbError(System.Data.OleDb.OleDbException ex, string prefix)
        {
            var sb = new System.Text.StringBuilder();
            if (ex.Errors != null && ex.Errors.Count > 0)
            {
                foreach (System.Data.OleDb.OleDbError e in ex.Errors)
                    sb.AppendLine($"- [{e.NativeError}] {e.Source} : {e.Message}");
            }
            else
            {
                sb.AppendLine(ex.Message);
            }
            if (ex.InnerException != null)
                sb.AppendLine($"Inner: {ex.InnerException.Message}");

            MessageBox.Show($"{prefix}\r\n{sb}");
        }

        private async Task SendToDbData(TaskKind taskKind)
        {
            var tag = (taskKind == TaskKind.Calibration) ? "Cal" : "Veri";

            try
            {
                if (!Settings.Instance.USE_MES) return;

                bool pass = (taskKind == TaskKind.Calibration)
                    ? (Cal_Result_flag == 1)
                    : (Veri_Result_flag == 1);

                await SaveDbResultAsync(taskKind, pass);   
                LogToUI(tag, "MES 저장 완료", 1);
            }
            catch (System.Data.OleDb.OleDbException ex)
            {
                LogOleDbError(ex, "MES 저장 실패");
            }
            catch (Exception ex)
            {
                LogToUI(tag, $"MES 저장 실패(예외) : {ex.Message}", 0);
            }
        }
        #endregion

        #region Helper
        private static bool TryParseAscii03(byte[] resp, out ushort[] regs)
        {
            regs = null;
            if (resp == null || resp.Length < 7) return false;

            string line = Encoding.ASCII.GetString(resp).Trim(); // ":...."
            if (!line.StartsWith(":")) return false;

            // ':' 제외, CRLF 제거된 HEX 본문
            string hex = line.Substring(1);
            if (hex.Length % 2 != 0) return false;

            byte[] raw = HexToBytes(hex); // [Slave][Func][ByteCount][Data...][LRC]
            if (raw == null || raw.Length < 5) return false;

            // LRC 검증
            byte lrc = raw[raw.Length - 1];
            byte calc = ComputeLRC(raw, 0, raw.Length - 1);
            if (lrc != calc) return false;

            // 예외 응답 체크 (func | 0x80)
            byte func = raw[1];
            if ((func & 0x80) != 0) return false;
            if (func != 0x03) return false;

            int bc = raw[2];
            if (raw.Length != 3 + bc + 1) return false; // [ID][FC][BC][Data...][LRC]

            // 레지스터 배열로 변환 (big-endian)
            regs = new ushort[bc / 2];
            for (int i = 0; i < regs.Length; i++)
            {
                int hi = raw[3 + i * 2];
                int lo = raw[4 + i * 2];
                regs[i] = (ushort)(hi | (lo << 8));
            }
            return true;
        }

        private static byte[] HexToBytes(string hex)
        {
            var bytes = new byte[hex.Length / 2];
            for (int i = 0; i < bytes.Length; i++)
            {
                int hi = Convert.ToInt32(hex.Substring(i * 2, 1), 16);
                int lo = Convert.ToInt32(hex.Substring(i * 2 + 1, 1), 16);
                bytes[i] = (byte)((hi << 4) | lo);
            }
            return bytes;
        }

        private static byte ComputeLRC(byte[] buf, int offset, int len)
        {
            int sum = 0; for (int i = 0; i < len; i++) sum += buf[offset + i];
            return (byte)((-sum) & 0xFF);
        }

        private void SafeSetText(TextBox tb, string text)
        {
            if (tb.InvokeRequired) tb.BeginInvoke(new Action(() => tb.Text = text));
            else tb.Text = text;
        }

        private static bool TryParseAsciiToBytes(byte[] asciiFrame, out byte[] bytes)
        {
            bytes = null;
            if (asciiFrame == null || asciiFrame.Length < 3) return false;

            // 1) 문자열 변환
            string s = System.Text.Encoding.ASCII.GetString(asciiFrame);

            // 2) 시작 콜론(:) 위치
            int colon = s.IndexOf(':');
            if (colon >= 0) s = s.Substring(colon + 1);

            // 3) 줄바꿈 제거
            while (s.Length > 0 && (s[s.Length - 1] == '\r' || s[s.Length - 1] == '\n'))
                s = s.Substring(0, s.Length - 1);

            // 4) 짝수 길이 체크
            if ((s.Length % 2) != 0) return false;

            int len = s.Length / 2;
            byte[] buf = new byte[len];
            for (int i = 0; i < len; i++)
            {
                int hi = HexVal(s[2 * i]);
                int lo = HexVal(s[2 * i + 1]);
                if (hi < 0 || lo < 0) return false;
                buf[i] = (byte)((hi << 4) | lo);
            }
            bytes = buf;
            return true;

            int HexVal(char c)
            {
                if (c >= '0' && c <= '9') return c - '0';
                if (c >= 'A' && c <= 'F') return c - 'A' + 10;
                if (c >= 'a' && c <= 'f') return c - 'a' + 10;
                return -1;
            }
        }

        static bool IsEchoAck(byte[] tx, byte[] rx, out string reason) //단일 Write일 때 응답검사
        {
            reason = null;
            if (rx == null || rx.Length == 0) { reason = "응답 없음"; return false; }

            // 1) 문자열 비교(간단 경로)
            string sTx = Encoding.ASCII.GetString(tx).Trim().ToUpperInvariant();
            string sRx = Encoding.ASCII.GetString(rx).Trim().ToUpperInvariant();
            if (sTx == sRx) return true;
            else return false;
        }

        private bool IsWriteMultiAck(byte[] rx, byte slave, ushort startAddr, ushort qty) //멀티 Write일 때 응답검사
        {
            byte[] frame;

            // Modbus ASCII라면 바이트 배열로 변환
            if (!TryParseAsciiToBytes(rx, out frame))
            {
                // 아니면 원본 그대로 사용(이미 바이너리일 수 있음)
                frame = rx;
            }

            // 최소 7바이트: [Slave][Func][AddrHi][AddrLo][QtyHi][QtyLo][LRC/CRC...]
            if (frame == null || frame.Length < 7) return false;

            // 프레임 어디에 있어도 패턴 스캔해서 확인 (헤더 붙어있어도 안전)
            for (int i = 0; i <= frame.Length - 6; i++)
            {
                if (frame[i] == slave && frame[i + 1] == 0x10)
                {
                    ushort addr = (ushort)((frame[i + 2] << 8) | frame[i + 3]);
                    ushort q = (ushort)((frame[i + 4] << 8) | frame[i + 5]);
                    return (addr == startAddr && q == qty);
                }
            }
            return false;
        }

        public static byte[] BuildProfileData(bool IsCal)
        {
            ushort[] words = new ushort[24];
            if (IsCal) //Calibration
            {
                // 온도 (0~11)
                words[0] = (ushort)Settings.Instance.Point1_Temp;
                words[1] = (ushort)Settings.Instance.Point2_Temp;
                words[2] = (ushort)Settings.Instance.Point3_Temp;
                words[3] = (ushort)Settings.Instance.Point4_Temp;
                words[4] = (ushort)Settings.Instance.Point5_Temp;
                words[5] = (ushort)Settings.Instance.Point6_Temp;
                words[6] = (ushort)Settings.Instance.Point7_Temp;
                words[7] = (ushort)Settings.Instance.Point8_Temp;
                words[8] = (ushort)Settings.Instance.Point9_Temp;
                words[9] = (ushort)Settings.Instance.Point10_Temp;
                words[10] = (ushort)Settings.Instance.Point11_Temp;
                words[11] = (ushort)Settings.Instance.Point12_Temp;

                // 시간 (12~23)
                words[12] = (ushort)Settings.Instance.Point1_Time;
                words[13] = (ushort)Settings.Instance.Point2_Time;
                words[14] = (ushort)Settings.Instance.Point3_Time;
                words[15] = (ushort)Settings.Instance.Point4_Time;
                words[16] = (ushort)Settings.Instance.Point5_Time;
                words[17] = (ushort)Settings.Instance.Point6_Time;
                words[18] = (ushort)Settings.Instance.Point7_Time;
                words[19] = (ushort)Settings.Instance.Point8_Time;
                words[20] = (ushort)Settings.Instance.Point9_Time;
                words[21] = (ushort)Settings.Instance.Point10_Time;
                words[22] = (ushort)Settings.Instance.Point11_Time;
                words[23] = (ushort)Settings.Instance.Point12_Time;
            }
            else //Verification
            {
                // 온도 (0~11)
                words[0] = (ushort)Settings.Instance.Point1_Temp_Verification;
                words[1] = (ushort)Settings.Instance.Point2_Temp_Verification;
                words[2] = (ushort)Settings.Instance.Point3_Temp_Verification;
                words[3] = (ushort)Settings.Instance.Point4_Temp_Verification;
                words[4] = (ushort)Settings.Instance.Point5_Temp_Verification;
                words[5] = (ushort)Settings.Instance.Point6_Temp_Verification;
                words[6] = (ushort)Settings.Instance.Point7_Temp_Verification;
                words[7] = (ushort)Settings.Instance.Point8_Temp_Verification;
                words[8] = (ushort)Settings.Instance.Point9_Temp_Verification;
                words[9] = (ushort)Settings.Instance.Point10_Temp_Verification;
                words[10] = (ushort)Settings.Instance.Point11_Temp_Verification;
                words[11] = (ushort)Settings.Instance.Point12_Temp_Verification;

                // 시간 (12~23)
                words[12] = (ushort)Settings.Instance.Point1_Time_Verification;
                words[13] = (ushort)Settings.Instance.Point2_Time_Verification;
                words[14] = (ushort)Settings.Instance.Point3_Time_Verification;
                words[15] = (ushort)Settings.Instance.Point4_Time_Verification;
                words[16] = (ushort)Settings.Instance.Point5_Time_Verification;
                words[17] = (ushort)Settings.Instance.Point6_Time_Verification;
                words[18] = (ushort)Settings.Instance.Point7_Time_Verification;
                words[19] = (ushort)Settings.Instance.Point8_Time_Verification;
                words[20] = (ushort)Settings.Instance.Point9_Time_Verification;
                words[21] = (ushort)Settings.Instance.Point10_Time_Verification;
                words[22] = (ushort)Settings.Instance.Point11_Time_Verification;
                words[23] = (ushort)Settings.Instance.Point12_Time_Verification;
            }

            // ushort → byte (big-endian) 변환
            byte[] bytes = new byte[words.Length * 2];
            for (int i = 0; i < words.Length; i++)
            {
                bytes[i * 2] = (byte)(words[i] >> 8);      // 상위 바이트
                bytes[i * 2 + 1] = (byte)(words[i] & 0xFF); // 하위 바이트
            }

            return bytes;
        }

        private static double TrimmedMean(List<double> data) //최대 최소 제외 평균구하는 함수
        {
            if (data == null || data.Count == 0) return double.NaN;
            if (data.Count <= 2) return data.Average();

            double sum = 0.0;
            double min = double.PositiveInfinity;
            double max = double.NegativeInfinity;

            for (int i = 0; i < data.Count; i++)
            {
                double v = data[i];
                sum += v;
                if (v < min) min = v;
                if (v > max) max = v;
            }
            return (sum - min - max) / (data.Count - 2);
        }

        private static string SafeAscii(byte[] buf)
        {
            try { return Encoding.ASCII.GetString(buf); }
            catch { return BitConverter.ToString(buf); }
        }

        private static string Bracket(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return name;
            var parts = name.Split('.');
            for (int i = 0; i < parts.Length; i++)
            {
                parts[i] = "[" + parts[i].Trim().Trim('[', ']') + "]";
            }
            return string.Join(".", parts);
        }

        private async Task SaveDbResultAsync(TaskKind kind, bool pass)
        {
            string cs = OleDbHelper.ConnStr.SqlServer(Settings.Instance.DB_IP,
                Settings.Instance.DB_PORT, Settings.Instance.DB_NAME, Settings.Instance.DB_USER, Settings.Instance.DB_PW);

            string tbl = Bracket(Settings.Instance.DB_TABLE);
            string sql = $@"INSERT INTO {tbl} ([TaskName], [Calibration Result], [Verification Result]) VALUES (?,?,?)";

            string taskName = (kind == TaskKind.Calibration) ? "Calibration" : "Verification";

            object calVal = (kind == TaskKind.Calibration) ? (object)(pass ? "PASS" : "NG") : (object)DBNull.Value;
            object verVal = (kind == TaskKind.Verification) ? (object)(pass ? "PASS" : "NG") : (object)DBNull.Value;

            await OleDbHelper.ExecuteNonQueryAsync(cs, sql, 30, CancellationToken.None,
                new OleDbParameter { OleDbType = OleDbType.VarChar, Value = taskName },
                new OleDbParameter { OleDbType = OleDbType.VarChar, IsNullable = true, Value = calVal },
                new OleDbParameter { OleDbType = OleDbType.VarChar, IsNullable = true, Value = verVal }
                );

        }
        #endregion

        #region Connect

        private async void btnConnect_Click_1(object sender, EventArgs e)
        {
            string portName = cboxDevCom.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(portName) || portName == "None") return;
            try
            {
                if (DevicePort != null && DevicePort.IsOpen)
                {
                    DevicePort.Disconnect();
                }

                DevicePort = new SerialChannelPort(1);
                DevicePort.Connect(portName, 115200);
                DevicePort.LogCommToUI = (box, s, tx) => LogCommToUI(tboxDeviceComm, s, tx);
                LogToUI("Device", $"디바이스 연결 성공");
                //AllClearDeviceManualLabel();

                var pkt = new CDCProtocol(Variable.SLAVE, Variable.WRITE, Variable.CHARGE_OFF).GetPacket();
                var resp = await DevicePort.SendAndReceivePacketAsync(pkt, 1500);

                if (resp == null)
                {
                    LogToUI("Device", "Charge OFF 명령 실패", 0);
                    return;
                }

                if (IsEchoAck(pkt, resp, out var whyNot))
                {
                    LogToUI("Device", "Charge OFF 적용 완료", 1);
                }
                else
                {
                    LogToUI("Device", $"Charge OFF 실패: {whyNot}", 0);
                }

                Console.WriteLine($"{BitConverter.ToString(resp)}");
                // ✅ MCU ID 읽기: Addr=0x0041, Qty=4
                await ReadAndShowMcuIdAsync();
                await ReadAndShowSerialNoAsync();
                await ReadAndShowFwVerAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine("디바이스 연결 실패" + ex.Message);
                LogToUI("Device", $"디바이스 연결 실패" + ex.Message);
            }
        }

        private void btnDisConnect_Click_1(object sender, EventArgs e)
        {
            if (DevicePort == null || !DevicePort.IsOpen) return;
            try
            {
                DevicePort.Disconnect();

                tboxDeviceId.Text = string.Empty;
                tboxSerialNo.Text = string.Empty;
                tboxFwVer.Text = string.Empty;
                LogToUI("Device", $"디바이스 해제 완료");
                //AllClearDeviceManualLabel();
            }
            catch (Exception ex)
            {
                Console.WriteLine("디바이스 해제 실패" + ex.Message);
                LogToUI("Device", "디바이스 해제 실패" + ex.Message);
            }
        }

        private void btnConnectMeas_Click(object sender, EventArgs e)
        {
            string portName = cboxMeasuCom.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(portName) || portName == "None") return;
            try
            {
                if (MeasPort != null && MeasPort.IsOpen)
                {
                    MeasPort.Disconnect();
                }

                MeasPort = new MeasuringChannelPort(1);
                MeasPort.Connect(portName);
                MeasPort.LogCommToUI = (box, s, tx) => LogCommToUI(tboxMeasComm, s, tx);
                LogToUI("Maurer", $"온도계 연결 성공");
                //AllClearMaurerManualLabel();
            }
            catch (Exception ex)
            {
                Console.WriteLine("온도계 연결 실패" + ex.Message);
                LogToUI("Maurer", $"온도계 연결 실패" + ex.Message);
            }
        }

        private void btnDisConnectMeas_Click(object sender, EventArgs e)
        {
            if (MeasPort == null || !MeasPort.IsOpen) return;
            try
            {
                MeasPort.Disconnect();

                LogToUI("Maurer", $"온도계 연결 해제 완료");
                //AllClearMaurerManualLabel();
            }
            catch (Exception ex)
            {
                Console.WriteLine("온도계 연결 해제 실패" + ex.Message);
                LogToUI("Maurer", "온도계 연결 해제 실패" + ex.Message);
            }
        }

        private void btnRescan_Click_1(object sender, EventArgs e)
        {
            RescanSerialPort();
        }

        private void RescanSerialPort()
        {
            try
            {
                // 현재 PC에 연결된 모든 COM 포트 가져오기
                string[] portNames = SerialPort.GetPortNames();

                ComboBox[] boxes = { cboxDevCom, cboxMeasuCom };

                foreach (var box in boxes)
                {
                    box.Items.Clear();
                    box.Items.Add("None");
                    box.Items.AddRange(portNames);
                    box.SelectedIndex = 0;
                }

                // 새로운 COM 포트 목록 추가

            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while searching for port " + ex.Message);
            }
        }

        private async Task ReadAndShowMcuIdAsync()
        {
            var pkt = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.MCU_ID).GetPacket();
            var resp = await DevicePort.SendAndReceivePacketAsync(pkt, 1500);

            if(!TryParseAscii03(resp, out var regs))
            {
                LogToUI("Device", "MCU ID 읽기 실패", 0);
                return;
            }

            // 표시 포맷 ①: 16비트 4개를 하이픈으로 (예: 7094-8C0E-2D4A-2FF5)
            string uidFmt1 = $"{regs[0]:X4}-{regs[1]:X4}-{regs[2]:X4}-{regs[3]:X4}";

            // 표시 포맷 ②: 32비트 2개로 (예: 70948C0E-2D4A2FF5)
            uint u0 = (regs[0]) | ((uint)regs[1] << 16);
            uint u1 = (regs[2]) | ((uint)regs[3] << 16);
            string uidFmt2 = $"{u1:X8}{u0:X8}";

            SafeSetText(tboxDeviceId, uidFmt2);  // 둘 중 원하는 걸로 선택
            LogToUI("Device", $"MCU ID = {uidFmt2}", 1);

        }

        private async Task ReadAndShowSerialNoAsync()
        {
            var pkt = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.SERIAL_NO).GetPacket();
            var resp = await DevicePort.SendAndReceivePacketAsync(pkt, Device_ReadTimeOut_Value);

            if (!TryParseAscii03(resp, out var regs))
            {
                SafeSetText(tboxSerialNo, "(SN 읽기 실패)");
                return;
            }

            // 16비트 레지스터 8개 -> 16바이트 -> ASCII
            var bytes = new byte[regs.Length * 2];
            for (int i = 0; i < regs.Length; i++)
            {
                bytes[2 * i] = (byte)(regs[i] >> 8);   // High
                bytes[2 * i + 1] = (byte)(regs[i] & 0xFF); // Low
            }

            string sn = Encoding.ASCII.GetString(bytes)
                            .TrimEnd('\0', ' '); // 뒤쪽 널/공백 제거 (필요 시)

            // 만약 인쇄 불가 문자가 섞여 있으면 보기 좋게 걸러서 대안 제공
            if (sn.Any(ch => ch < 0x20 || ch > 0x7E))
            {
                sn = string.Concat(bytes.Select(b => (b >= 0x20 && b <= 0x7E) ? (char)b : '\0'))
                         .TrimEnd('\0');
                if (string.IsNullOrWhiteSpace(sn))
                    sn = string.Join("", regs.Select(r => r.ToString("X4"))); // 완전 fallback (HEX)
            }

            SafeSetText(tboxSerialNo, sn);
            LogToUI("Device", $"SN = {sn}", 1);
        }

        private async Task ReadAndShowFwVerAsync()
        {
            var pkt = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.FW_VER).GetPacket();
            var resp = await DevicePort.SendAndReceivePacketAsync(pkt, Device_ReadTimeOut_Value);

            if (!TryParseAscii03(resp, out var regs) || regs.Length != 2)
            {
                SafeSetText(tboxFwVer, "(FW_VER 읽기 실패)");
                return;
            }

            // 레지스터 2개(각 2바이트) → 총 4바이트 → ASCII 숫자
            byte majorHi = (byte)(regs[0] >> 8);
            byte majorLo = (byte)(regs[0] & 0xFF);
            byte minorHi = (byte)(regs[1] >> 8);
            byte minorLo = (byte)(regs[1] & 0xFF);

            // 기대값: 모두 '0'~'9'(0x30~0x39)
            string major = new string(new[] { (char)majorHi, (char)majorLo });
            string minor = new string(new[] { (char)minorHi, (char)minorLo });

            // 혹시 숫자 아닌 값이 섞여 있으면 안전하게 숫자만 추림
            if (!(major.All(ch => ch >= '0' && ch <= '9'))) major = $"{regs[0] & 0xFF:D2}";
            if (!(minor.All(ch => ch >= '0' && ch <= '9'))) minor = $"{regs[1] & 0xFF:D2}";

            string ver = $"v{major}.{minor}";
            SafeSetText(tboxFwVer, ver);
            LogToUI("Device", $"FW_VER = {ver}", 1);
        }



        #endregion

        #region Profile
        private void btnOpenProfileSetting_Click(object sender, EventArgs e)
        {
            ProfileSettingForm profile_form = new ProfileSettingForm(this);
            profile_form.Show();
        }
    

        private void btnProfileSelect_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "JSON files (*.json)|*.json";
            openFileDialog.Title = "프로파일 열기";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                string fileName = System.IO.Path.GetFileName(filePath);

                SettingNowProfile(filePath, fileName);
            }
        }


        private void SettingNowProfile(string filePath, string fileName)
        {
            try
            {
                string json = File.ReadAllText(filePath);
                dynamic profile = JsonConvert.DeserializeObject(json);

                // Type 검사
                if (profile["Type"]?.ToString() != "Profile")
                {
                    MessageBox.Show("유효하지 않은 프로파일 형식입니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                JObject settings = profile["Settings"]?.ToObject<JObject>();

                if (settings == null)
                {
                    MessageBox.Show("레시피 파일에 설정값이 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                tboxFilePath.Text = filePath;
                Settings.Instance.Profile_File_Path = filePath;
                tboxProfile.Text = fileName;
                Settings.Instance.Profile_File = fileName;

                ApplySettingsFromJson(settings); //파일 불러와서 저장값에 적용

                dgViewNowProfile.DataSource = null;
                dgViewNowProfile.DataSource = CreateNowProfileTable(); //저장값에 따라 테이블 만들기

                Settings.Instance.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Exception occured: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }

        private void ApplySettingsFromJson(JObject settings)
        {
            //온도
            Settings.Instance.Point1_Temp = settings["Point1_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point2_Temp = settings["Point2_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point3_Temp = settings["Point3_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point4_Temp = settings["Point4_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point5_Temp = settings["Point5_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point6_Temp = settings["Point6_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point7_Temp = settings["Point7_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point8_Temp = settings["Point8_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point9_Temp = settings["Point9_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point10_Temp = settings["Point10_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point11_Temp = settings["Point11_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point12_Temp = settings["Point12_Temp"]?.ToObject<int>() ?? 0;
            //시간
            Settings.Instance.Point1_Time = settings["Point1_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point2_Time = settings["Point2_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point3_Time = settings["Point3_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point4_Time = settings["Point4_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point5_Time = settings["Point5_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point6_Time = settings["Point6_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point7_Time = settings["Point7_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point8_Time = settings["Point8_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point9_Time = settings["Point9_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point10_Time = settings["Point10_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point11_Time = settings["Point11_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point12_Time = settings["Point12_Time"]?.ToObject<int>() ?? 0;

            Settings.Instance.Save();
        }

        private DataTable CreateNowProfileTable()
        {
            DataTable table = new DataTable();

            table.Columns.Add("구간", typeof(string));
            table.Columns.Add("1", typeof(int));
            table.Columns.Add("2", typeof(int));
            table.Columns.Add("3", typeof(int));
            table.Columns.Add("4", typeof(int));
            table.Columns.Add("5", typeof(int));
            table.Columns.Add("6", typeof(int));
            table.Columns.Add("7", typeof(int));
            table.Columns.Add("8", typeof(int));
            table.Columns.Add("9", typeof(int));
            table.Columns.Add("10", typeof(int));
            table.Columns.Add("11", typeof(int));
            table.Columns.Add("12", typeof(int));

            table.Rows.Add("온도", Settings.Instance.Point1_Temp, Settings.Instance.Point2_Temp, Settings.Instance.Point3_Temp, Settings.Instance.Point4_Temp,
                Settings.Instance.Point5_Temp, Settings.Instance.Point6_Temp, Settings.Instance.Point7_Temp, Settings.Instance.Point8_Temp,
                Settings.Instance.Point9_Temp, Settings.Instance.Point10_Temp, Settings.Instance.Point11_Temp, Settings.Instance.Point12_Temp);
            table.Rows.Add("시간", Settings.Instance.Point1_Time, Settings.Instance.Point2_Time, Settings.Instance.Point3_Time, Settings.Instance.Point4_Time,
                Settings.Instance.Point5_Time, Settings.Instance.Point6_Time, Settings.Instance.Point7_Time, Settings.Instance.Point8_Time,
                Settings.Instance.Point9_Time, Settings.Instance.Point10_Time, Settings.Instance.Point11_Time, Settings.Instance.Point12_Time);

            return table;
        }

        private void btnVerificationProfileSelect_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "JSON files (*.json)|*.json";
            openFileDialog.Title = "프로파일 열기";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                string fileName = System.IO.Path.GetFileName(filePath);

                SettingVerificationProfile(filePath, fileName);
            }
        }

        private void SettingVerificationProfile(string filePath, string fileName)
        {
            try
            {
                string json = File.ReadAllText(filePath);
                dynamic profile = JsonConvert.DeserializeObject(json);

                // Type 검사
                if (profile["Type"]?.ToString() != "Profile")
                {
                    MessageBox.Show("유효하지 않은 프로파일 형식입니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                JObject settings = profile["Settings"]?.ToObject<JObject>();

                if (settings == null)
                {
                    MessageBox.Show("레시피 파일에 설정값이 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                tboxFilePath_Verification.Text = filePath;
                Settings.Instance.Profile_File_Path_Verification = filePath;
                tboxProfile_Verification.Text = fileName;
                Settings.Instance.Profile_File_Verification = fileName;

                ApplySettingsFromJson_Verification(settings); //파일 불러와서 저장값에 적용

                dgViewVerificationProfile.DataSource = null;
                dgViewVerificationProfile.DataSource = CreateVerificationProfileTable(); //저장값에 따라 테이블 만들기

                Settings.Instance.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Exception occured: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }

        private DataTable CreateVerificationProfileTable()
        {
            DataTable table = new DataTable();

            table.Columns.Add("구간", typeof(string));
            table.Columns.Add("1", typeof(int));
            table.Columns.Add("2", typeof(int));
            table.Columns.Add("3", typeof(int));
            table.Columns.Add("4", typeof(int));
            table.Columns.Add("5", typeof(int));
            table.Columns.Add("6", typeof(int));
            table.Columns.Add("7", typeof(int));
            table.Columns.Add("8", typeof(int));
            table.Columns.Add("9", typeof(int));
            table.Columns.Add("10", typeof(int));
            table.Columns.Add("11", typeof(int));
            table.Columns.Add("12", typeof(int));

            table.Rows.Add("온도", Settings.Instance.Point1_Temp_Verification, Settings.Instance.Point2_Temp_Verification, Settings.Instance.Point3_Temp_Verification,
                Settings.Instance.Point4_Temp_Verification, Settings.Instance.Point5_Temp_Verification, Settings.Instance.Point6_Temp_Verification, Settings.Instance.Point7_Temp_Verification,
                Settings.Instance.Point8_Temp_Verification, Settings.Instance.Point9_Temp_Verification, Settings.Instance.Point10_Temp_Verification,
                Settings.Instance.Point11_Temp_Verification, Settings.Instance.Point12_Temp_Verification);
            table.Rows.Add("시간", Settings.Instance.Point1_Time_Verification, Settings.Instance.Point2_Time_Verification, Settings.Instance.Point3_Time_Verification, 
                Settings.Instance.Point4_Time_Verification, Settings.Instance.Point5_Time_Verification, Settings.Instance.Point6_Time_Verification, Settings.Instance.Point7_Time_Verification,
                Settings.Instance.Point8_Time_Verification, Settings.Instance.Point9_Time_Verification, Settings.Instance.Point10_Time_Verification,
                Settings.Instance.Point11_Time_Verification, Settings.Instance.Point12_Time_Verification);

            return table;
        }

        private void ApplySettingsFromJson_Verification(JObject settings)
        {
            //온도
            Settings.Instance.Point1_Temp_Verification = settings["Point1_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point2_Temp_Verification = settings["Point2_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point3_Temp_Verification = settings["Point3_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point4_Temp_Verification = settings["Point4_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point5_Temp_Verification = settings["Point5_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point6_Temp_Verification = settings["Point6_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point7_Temp_Verification = settings["Point7_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point8_Temp_Verification = settings["Point8_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point9_Temp_Verification = settings["Point9_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point10_Temp_Verification = settings["Point10_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point11_Temp_Verification = settings["Point11_Temp"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point12_Temp_Verification = settings["Point12_Temp"]?.ToObject<int>() ?? 0;
            //시간
            Settings.Instance.Point1_Time_Verification = settings["Point1_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point2_Time_Verification = settings["Point2_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point3_Time_Verification = settings["Point3_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point4_Time_Verification = settings["Point4_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point5_Time_Verification = settings["Point5_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point6_Time_Verification = settings["Point6_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point7_Time_Verification = settings["Point7_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point8_Time_Verification = settings["Point8_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point9_Time_Verification = settings["Point9_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point10_Time_Verification = settings["Point10_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point11_Time_Verification = settings["Point11_Time"]?.ToObject<int>() ?? 0;
            Settings.Instance.Point12_Time_Verification = settings["Point12_Time"]?.ToObject<int>() ?? 0;

            Settings.Instance.Save();
        }
        #endregion

        #region Device Manual
        private async Task Charge_Off_Manual()
        {
            int timeout = Device_ReadTimeOut_Value;
            // 1) 패킷 생성
            byte[] tx = new CDCProtocol(Variable.SLAVE, Variable.WRITE, Variable.CHARGE_OFF).GetPacket();

            // 2) 송수신
            byte[] rx = await DevicePort.SendAndReceivePacketAsync(tx, timeout);
            if (rx == null)
            {
                LogToUI("Device", "응답 없음", 0);
                return;
            }

            // 3) 응답 검증 및 로그
            string why;
            if (IsEchoAck(tx, rx, out why))
                LogToUI("Device", "Charge OFF 적용 완료", 1);
            else
                LogToUI("Device", "Charge OFF 적용 실패", 0);

            return;
        }
        private async Task Set_Init_R_Init_Temp_Manual()
        {
            int timeout = Device_ReadTimeOut_Value;
            // 1) 패킷 생성
            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.MULTI_WRITE, Variable.SET_INIT_R_INIT_TEMP).GetPacket();

            // 2) 송수신
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            if (rx1 == null || !IsWriteMultiAck(rx1, Variable.SLAVE, 0x00AF, 0x0002))
            {
                LogToUI("Device", "Set init R init temp 실패(멀티 쓰기 ACK 불일치)", 0);
                return;
            }

            byte[] tx2 = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.CHECK_INIT_R_INIT_TEMP).GetPacket();
            byte[] rx2 = await DevicePort.SendAndReceivePacketAsync(tx2, timeout);
            if (rx2 == null)
            {
                LogToUI("Device", "Check init R init Temp 응답 없음", 0);
                return;
            }

            byte[] frame;
            if (!TryParseAsciiToBytes(rx2, out frame))
                frame = rx2;

            if (frame.Length < 3)
            {
                LogToUI("Device", "Check init R init Temp 실패 (응답 프레임 짧음)", 0);
                return;
            }

            int byteCount = frame[2];
            int payloadOffset = 3;

            if (frame.Length < payloadOffset + byteCount)
            {
                LogToUI("Device", "Check init R init Temp 실패 (응답 ByteCount 불일치)", 0);
                return;
            }

            if (byteCount != 4) //레지스터 2개 == 4바이트
            {
                LogToUI("Device", $"Check Profile 실패 (예상 ByteCount=4, 수신={byteCount})", 0);
                return;
            }

            byte[] Byte_Heater_Init_Resistance = { frame[3], frame[4] };
            byte[] Byte_Heater_Init_Temperature = { frame[5], frame[6] };
            short Heater_Init_Resistance = BitConverter.ToInt16(Byte_Heater_Init_Resistance, 0);
            short Heater_Init_Temperature = BitConverter.ToInt16(Byte_Heater_Init_Temperature, 0);
            //byte[] Byte_expected_Resistance = { 0x00, 0x00 };
            byte[] Byte_expected_Temperature = { Variable.SET_INIT_R_INIT_TEMP[7], Variable.SET_INIT_R_INIT_TEMP[8] };
            //Array.Reverse(Byte_expected_Resistance);
            Array.Reverse(Byte_expected_Temperature);
            //short expected_Heater_Init_Resistance = BitConverter.ToInt16(Byte_expected_Resistance, 0);
            short expected_Heater_Init_Temperature = BitConverter.ToInt16(Byte_expected_Temperature, 0);

            short Resistance_Min = 0;
            short Resistance_Max = 2000;

            bool isResisOk = Heater_Init_Resistance > Resistance_Min && Heater_Init_Resistance < Resistance_Max;
            bool isTempOk = Heater_Init_Temperature == expected_Heater_Init_Temperature;

            if (!isResisOk || !isTempOk)
            {
                LogToUI("Device", "Set init R init Temp 실패", 0);
                LogToUI("Device", $"Resistance ==> Set = {Heater_Init_Resistance} [{Resistance_Min} ~ {Resistance_Max}]", isResisOk ? 1 : 0);
                LogToUI("Device", $"Temperature ==> Set = {Heater_Init_Temperature} [{expected_Heater_Init_Temperature}]", isTempOk ? 1 : 0);
                return;
            }

            LogToUI("Device", "Set init R init Temp 완료", 1);
            LogToUI("Device", $"Set = {Heater_Init_Resistance} [{Resistance_Min} ~ {Resistance_Max}]", isResisOk ? 1 : 0);
            LogToUI("Device", $"Set = {Heater_Init_Temperature} [{expected_Heater_Init_Temperature}]", isTempOk ? 1 : 0);

            return;
        }
        private async Task Set_Calibration_Profile_Manual()
        {
            if (string.IsNullOrEmpty(tboxProfile.Text) || string.IsNullOrEmpty(tboxFilePath.Text))
            {
                MessageBox.Show("프로파일을 선택해주세요.");
                return;
            }
            int timeout = Device_ReadTimeOut_Value;
            // 1) Settings.Instance 값으로 데이터 빌드 (온도12 + 시간12 → 총 24워드)
            byte[] data = BuildProfileData(true);
            byte[] payload = new byte[2 + 2 + 1 + data.Length];

            payload[0] = 0x00; payload[1] = 0x55;
            payload[2] = 0x00; payload[3] = 0x18;
            payload[4] = (byte)data.Length;

            Buffer.BlockCopy(data, 0, payload, 5, data.Length);

            // 2) 0x55 ~ 0x6C 연속 쓰기 (Write Multiple Registers)
            //    ↓ CDCProtocol 생성자는 네 프로젝트 규약에 맞게 수정해
            //    예시: (slave, func: WRITE_MULTI, startAddr, payload)
            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.MULTI_WRITE, payload).GetPacket();
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            if (rx1 == null || !IsWriteMultiAck(rx1, Variable.SLAVE, 0x0055, 0x0018))
            {
                LogToUI("Device", "프로파일 전송 실패(멀티 쓰기 ACK 불일치)", 0);
                return;
            }

            // 3) 0x6D에 0x0101 쓰기 (프로파일 인덱스/적용 트리거)
            //    예: (slave, func: WRITE, address, data2바이트)
            byte[] applyData = new byte[4] { 0x00, 0x6D, 0x01, 0x01 }; // 0x0101 (big-endian 전송이면 CDCProtocol에서 처리)
            byte[] tx2 = new CDCProtocol(Variable.SLAVE, Variable.WRITE, applyData).GetPacket();
            byte[] rx2 = await DevicePort.SendAndReceivePacketAsync(tx2, timeout);
            if (rx2 == null)
            {
                LogToUI("Device", "프로파일 적용 응답 없음", 0);
                return;
            }

            string why2;
            if (IsEchoAck(tx2, rx2, out why2))
                LogToUI("Device", "Set Calibration Profile 완료", 1);
            else
                LogToUI("Device", "Set Calibration Profile 실패", 0);

            return;
        }
        private async Task Check_Profile_Manual()
        {
            int timeout = Device_ReadTimeOut_Value;
            byte[] applyData = new byte[4] { 0x00, 0x6D, 0x00, 0x01 };
            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.WRITE, applyData).GetPacket();
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            string why1;
            if (rx1 == null || !IsEchoAck(tx1, rx1, out why1))
            {
                LogToUI("Device", "Check Profile 실패", 0);
                return;
            }

            byte[] tx2 = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.CHECK_PROFILE).GetPacket();
            byte[] rx2 = await DevicePort.SendAndReceivePacketAsync(tx2, timeout);
            if (rx2 == null)
            {
                LogToUI("Device", "Check Profile 실패 (프로파일 Read 응답 없음)", 0);
                return;
            }

            byte[] frame;
            if (!TryParseAsciiToBytes(rx2, out frame))
                frame = rx2;

            if (frame.Length < 3)
            {
                LogToUI("Device", "Check Profile 실패 (응답 프레임 짧음)", 0);
                return;
            }

            int byteCount = frame[2];
            int payloadOffset = 3;

            if (frame.Length < payloadOffset + byteCount)
            {
                LogToUI("Device", "Check Profile 실패 (응답 ByteCount 불일치)", 0);
                return;
            }

            if (byteCount != 48) // 24워드 * 2바이트
            {
                LogToUI("Device", $"Check Profile 실패 (예상 ByteCount=48, 수신={byteCount})", 0);
                return;
            }

            // 기대값 (big-endian → WORD)
            byte[] expected = BuildProfileData(true); // 48 bytes
            ushort[] expectedWords = new ushort[24];
            for (int i = 0; i < 24; i++)
            {
                expectedWords[i] = (ushort)((expected[i * 2] << 8) | expected[i * 2 + 1]);
            }

            // 응답값 (장치가 lo,hi 순서로 보냄 → 리틀엔디언 스왑)
            ushort[] receivedWords = new ushort[24];
            for (int i = 0; i < 24; i++)
            {
                int idx = payloadOffset + i * 2;
                receivedWords[i] = (ushort)(frame[idx] | (frame[idx + 1] << 8)); // lo | (hi<<8)
            }

            bool same = true;
            string tempLog = "온도 값 비교:\r\n";
            string timeLog = "시간 값 비교:\r\n";

            for (int i = 0; i < 12; i++)
            {
                ushort expT = expectedWords[i];        // 기대 온도 i+1
                ushort expTi = expectedWords[12 + i];   // 기대 시간 i+1

                ushort actT = receivedWords[i * 2];     // 읽은 온도 i+1 (짝수 인덱스)
                ushort actTi = receivedWords[i * 2 + 1]; // 읽은 시간 i+1 (홀수 인덱스)

                tempLog += string.Format("Point{0}_Temp: 설정={1}, 읽음={2}\r\n", i + 1, expT, actT);
                timeLog += string.Format("Point{0}_Time: 설정={1}, 읽음={2}\r\n", i + 1, expTi, actTi);

                if (expT != actT || expTi != actTi)
                    same = false;
            }

            LogToUI("Device",
                (same ? "Check Profile OK\r\n" : "Check Profile 불일치\r\n") + tempLog + timeLog,
                same ? 1 : 0);

            return;
        }
        private async Task Set_TR_Mode_Off_Manual()
        {
            int timeout = Device_ReadTimeOut_Value;
            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.MULTI_WRITE, Variable.SET_TR_MODE_OFF).GetPacket();
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            if (rx1 == null || !IsWriteMultiAck(rx1, Variable.SLAVE, 0x0055, 0x000C))
            {
                LogToUI("Device", "TR MODE OFF 실패(멀티 쓰기 ACK 불일치)", 0);
                return;
            }

            byte[] applyData = new byte[4] { 0x00, 0x6D, 0x01, 0x00 };
            byte[] tx2 = new CDCProtocol(Variable.SLAVE, Variable.WRITE, applyData).GetPacket();
            byte[] rx2 = await DevicePort.SendAndReceivePacketAsync(tx2, timeout);
            if (rx2 == null)
            {
                LogToUI("Device", "데이터 쓰기 응답 없음", 0);
                return;
            }

            string why2;
            if (!IsEchoAck(tx2, rx2, out why2))
            {
                LogToUI("Device", "Set TR Mode Off 실패", 0);
                return;
            }



            applyData = new byte[4] { 0x00, 0x6D, 0x00, 0x00 }; //0x6D에 0x0000 write
            byte[] tx3 = new CDCProtocol(Variable.SLAVE, Variable.WRITE, applyData).GetPacket();
            byte[] rx3 = await DevicePort.SendAndReceivePacketAsync(tx3, timeout);
            string why3;
            if (rx3 == null || !IsEchoAck(tx3, rx3, out why3))
            {
                LogToUI("Device", "Check Profile 실패", 0);
                return;
            }

            byte[] tx4 = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.CHECK_TR_MODE).GetPacket();
            byte[] rx4 = await DevicePort.SendAndReceivePacketAsync(tx4, timeout);
            if (rx4 == null)
            {
                LogToUI("Device", "Check TR Mode 실패 (Read 응답 없음)", 0);
                return;
            }

            byte[] frame;
            if (!TryParseAsciiToBytes(rx4, out frame)) frame = rx4;

            if (frame.Length < 3)
            {
                LogToUI("Device", "Check TR Mode 실패 (응답 프레임 짧음)", 0);
                return;
            }

            int byteCount = frame[2];
            int payloadOffset = 3;
            if (byteCount != 24 || frame.Length < payloadOffset + byteCount)
            {
                LogToUI("Device", $"Check TR Mode 실패 (ByteCount=24 예상, 수신={byteCount})", 0);
                return;
            }

            ushort[] temps = new ushort[12];
            for (int i = 0; i < 12; i++)
            {
                int idx = payloadOffset + i * 2;
                temps[i] = (ushort)(frame[idx] | (frame[idx + 1] << 8));
            }


            ushort[] expTemps = new ushort[12]; expTemps[3] = 30;

            bool same = true;
            var sb = new StringBuilder();
            sb.AppendLine("TR Mode Off 값 비교");
            for (int i = 0; i < 12; i++)
            {
                sb.AppendFormat("Point{0}_Temp: 설정={1}, 읽음={2}\r\n", i + 1, expTemps[i], temps[i]);
                if (expTemps[i] != temps[i]) same = false;
            }

            LogToUI("Device", (same ? "Check TR Mode OK\n" : "Check TR Mode 불일치\n") + sb.ToString(), same ? 1 : 0);

            return;
        }
        private async Task ABC_Value_Reset_Manual()
        {
            int timeout = Device_ReadTimeOut_Value;

            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.MULTI_WRITE, Variable.ABC_RESET_WRITE).GetPacket();
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            if (rx1 == null || !IsWriteMultiAck(rx1, Variable.SLAVE, 0x000E, 0x0006))
            {
                LogToUI("Device", "ABC Value Reset 실패 (멀티쓰기 ACK 불일치)", 0);
                return;
            }

            byte[] tx2 = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.ABC_RESET_READ).GetPacket();
            byte[] rx2 = await DevicePort.SendAndReceivePacketAsync(tx2, timeout);
            if (rx2 == null)
            {
                LogToUI("Device", "ABC Value Reset 실패 (Read 응답 없음)", 0);
                return;
            }

            byte[] frame;
            if (!TryParseAsciiToBytes(rx2, out frame)) frame = rx2;


            if (frame.Length < 3)
            {
                LogToUI("Device", "ABC Value Reset 실패 (응답 프레임 짧음)", 0);
                return;
            }

            int payloadOffset = 3;
            int byteCount = frame[2];
            if (byteCount != 12 || frame.Length < payloadOffset + byteCount)
            {
                LogToUI("Device", $"ABC Value Reset 실패 (ByteCount=12 예상, 수신={byteCount})", 0);
                return;
            }

            byte[] Byte_A = { frame[4], frame[3], frame[6], frame[5] };
            byte[] Byte_B = { frame[8], frame[7], frame[10], frame[9] };
            byte[] Byte_C = { frame[12], frame[11], frame[14], frame[13] };
            float A = BitConverter.ToSingle(Byte_A, 0);
            float B = BitConverter.ToSingle(Byte_B, 0);
            float C = BitConverter.ToSingle(Byte_C, 0);


            byte[] Byte_expected_A = { Variable.ABC_RESET_WRITE[5], Variable.ABC_RESET_WRITE[6], Variable.ABC_RESET_WRITE[7], Variable.ABC_RESET_WRITE[8] };
            byte[] Byte_expected_B = { Variable.ABC_RESET_WRITE[9], Variable.ABC_RESET_WRITE[10], Variable.ABC_RESET_WRITE[11], Variable.ABC_RESET_WRITE[12] };
            byte[] Byte_expected_C = { Variable.ABC_RESET_WRITE[13], Variable.ABC_RESET_WRITE[14], Variable.ABC_RESET_WRITE[15], Variable.ABC_RESET_WRITE[16] };

            float expected_A = BitConverter.ToSingle(Byte_expected_A, 0);
            float expected_B = BitConverter.ToSingle(Byte_expected_B, 0);
            float expected_C = BitConverter.ToSingle(Byte_expected_C, 0);

            Console.WriteLine($"return : {BitConverter.ToString(Byte_A)} expected : {BitConverter.ToString(Byte_expected_A)}");
            Console.WriteLine($"return : {BitConverter.ToString(Byte_B)} expected : {BitConverter.ToString(Byte_expected_B)}");
            Console.WriteLine($"return : {BitConverter.ToString(Byte_C)} expected : {BitConverter.ToString(Byte_expected_C)}");

            bool IsAOk = A == expected_A;
            bool IsBOk = B == expected_B;
            bool IsCOk = C == expected_C;

            if (!IsAOk || !IsBOk || !IsCOk)
            {
                LogToUI("Device", "ABC Value Reset Fail", 0);
                LogToUI("Device", $"A : 응답 = {A} 기대 = {expected_A} B : 응답 = {B} 기대 = {expected_B} C : 응답 = {C} 기대 = {expected_C}", 0);
                return;
            }

            LogToUI("Device", "ABC Value Reset Success", 1);
            LogToUI("Device", $"A : 응답 = {A} 기대 = {expected_A} B : 응답 = {B} 기대 = {expected_B} C : 응답 = {C} 기대 = {expected_C}", 1);

            return;
        }
        private async Task Select_Use_Profile_Manual()
        {
            int timeout = Device_ReadTimeOut_Value;
            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.WRITE, Variable.SELECT_USE_PROFILE_WRITE).GetPacket();
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            string why1;
            if (rx1 == null || !IsEchoAck(tx1, rx1, out why1))
            {
                LogToUI("Device", "Select Use Profile Write 실패", 0);
                return;
            }

            byte[] tx2 = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.SELECT_USE_PROFILE_READ).GetPacket();
            byte[] rx2 = await DevicePort.SendAndReceivePacketAsync(tx2, timeout);

            if (rx2 == null)
            {
                LogToUI("Device", "Select Use Profile Read 실패", 0);
                return;
            }

            byte[] frame;
            if (!TryParseAsciiToBytes(rx2, out frame)) frame = rx2;

            if (frame.Length < 3)
            {
                LogToUI("Device", "Select Use Profile Read 실패 (응답 프레임 짧음)", 0);
                return;
            }

            int payloadOffset = 3;
            int byteCount = frame[2];
            if (byteCount != 2 || frame.Length < payloadOffset + byteCount)
            {
                LogToUI("Device", $"Select Use Profile Read 실패 (ByteCount=2 예상, 수신={byteCount})", 0);
                return;
            }
            //write 빅엔디안 read 리틀엔디안?
            byte[] Byte_Profile_Select_Index = { frame[payloadOffset], frame[payloadOffset + 1] };
            ushort Profile_Select_Index = BitConverter.ToUInt16(Byte_Profile_Select_Index, 0); //받는거 리틀엔디안
            byte[] Byte_expected_Result = { Variable.SELECT_USE_PROFILE_WRITE[2], Variable.SELECT_USE_PROFILE_WRITE[3] };
            Array.Reverse(Byte_expected_Result); //주는거 빅엔디안
            ushort expected_Result = BitConverter.ToUInt16(Byte_expected_Result, 0);


            bool IsOk = Profile_Select_Index == expected_Result;

            if (!IsOk)
            {
                LogToUI("Device", $"Select Use Profile 실패", 0);
                LogToUI("Device", $"적용 : {Profile_Select_Index} [{expected_Result}]", 0);
                return;
            }

            LogToUI("Device", $"Select Use Profile 성공", 1);
            LogToUI("Device", $"적용 : {Profile_Select_Index} [{expected_Result}]", 1);

            return;
        }
        private async Task Heating_Start_Manual()
        {
            int timeout = Device_ReadTimeOut_Value;
            // 1) 패킷 생성
            byte[] tx = new CDCProtocol(Variable.SLAVE, Variable.WRITE, Variable.HEATING_START).GetPacket();

            // 2) 송수신
            byte[] rx = await DevicePort.SendAndReceivePacketAsync(tx, timeout);
            if (rx == null)
            {
                LogToUI("Device", "응답 없음", 0);
                return;
            }

            // 3) 응답 검증 및 로그
            string why;
            if (IsEchoAck(tx, rx, out why))
                LogToUI("Device", "Heating Start 적용 완료", 1);
            else
                LogToUI("Device", "Heating Start 적용 실패", 0);

            return;
        }
        private async Task Heating_Stop_Manual()
        {
            int timeout = Device_ReadTimeOut_Value;
            // 1) 패킷 생성
            byte[] tx = new CDCProtocol(Variable.SLAVE, Variable.WRITE, Variable.HEATING_STOP).GetPacket();

            // 2) 송수신
            byte[] rx = await DevicePort.SendAndReceivePacketAsync(tx, timeout);
            if (rx == null)
            {
                LogToUI("Device", "응답 없음", 0);
                return;
            }

            // 3) 응답 검증 및 로그
            string why;
            if (IsEchoAck(tx, rx, out why))
                LogToUI("Device", "Heating Stop 적용 완료", 1);
            else
                LogToUI("Device", "Heating Stop 적용 실패", 0);

            return;
        }
        private async Task Get_Device_Data_Manual()
        {
            int timeout = 90;
            byte[] tx = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.GET_DEVICE_DATA).GetPacket();
            byte[] rx = await DevicePort.SendAndReceivePacketAsync(tx, timeout);
            if (rx == null)
            {
                LogToUI("Device", "Device Data 응답 없음", 0);
                return;
            }

            byte[] frame;
            if (!TryParseAsciiToBytes(rx, out frame)) frame = rx;


            if (frame.Length < 3)
            {
                LogToUI("Device", "ABC Value Reset 실패 (응답 프레임 짧음)", 0);
                return;
            }

            int payloadOffset = 3;
            int byteCount = frame[2];
            if (byteCount != 26 || frame.Length < payloadOffset + byteCount)
            {
                LogToUI("Device", $"Get Device Data 실패 (ByteCount=26 예상, 수신={byteCount})", 0);
                return;
            }

            ushort U16(int regIndex)
            {
                int i = payloadOffset + regIndex * 2;
                return (ushort)(frame[i] | (frame[i + 1]) << 8); // Lo,Hi → U16
            }

            var status = U16(0);         // enum
            var stickStatus = U16(1);         // enum
            var profileStep = U16(2);         // enum

            var runTimeSec = U16(3) / 10.0;  // x10 0.0 [sec]

            var pvSusTempP_C = U16(4) / 10.0 - 100.0; // x10, +100.00 [°C]
            var pvSusTempA_C = U16(5) / 10.0 - 100.0; // x10, +100.00 [°C]

            var heaterCurrent_mA = U16(6);       // mA (스케일 표기 없음 → 그대로)
            var heaterVolt_mV = U16(7);       // mV (스케일 표기 없음 → 그대로)

            var pvPcbTemp_C = U16(8) / 10.0 + 50.0;  // (+50.0 [°C] 표기 → x10 - 50.0 추정)

            var pvDuty_pct = U16(9);         // 문서에 스케일 표기 '0 [%]' → 그대로(필요시 /10 적용)

            var battVolt_mV = U16(10);        // mV
            var battTemp_C = U16(11) / 10.0 + 50.0; // x10, +50.00 [°C]
            var battSoC_pct = U16(12) / 10.0; // x10 0.0 [%]

            if(InvokeRequired)
            {
                BeginInvoke((Action)(() =>
                {
                    lblStatus_Manual.Text = status.ToString();
                    lblStickStatus_Manual.Text = stickStatus.ToString();
                    lblProfileStep_Manual.Text = profileStep.ToString();
                    lblRunTime_Manual.Text = runTimeSec.ToString();
                    lblPVSUSTempP_Manual.Text = pvSusTempP_C.ToString();
                    lblPVSUSTempA_Manual.Text = pvSusTempA_C.ToString();
                    lblHeaterCurrent_Manual.Text = heaterCurrent_mA.ToString();
                    lblHeaterVolt_Manual.Text = heaterVolt_mV.ToString();
                    lblPVPCBTemp_Manual.Text = pvPcbTemp_C.ToString();
                    lblPVDuty_Manual.Text = pvDuty_pct.ToString();
                    lblBATTVolt_Manual.Text = battVolt_mV.ToString();
                    lblBATTTemp_Manual.Text = battTemp_C.ToString();
                    lblBATTSoc_Manual.Text = battSoC_pct.ToString();
                }));
            }
            else
            {
                lblStatus_Manual.Text = status.ToString();
                lblStickStatus_Manual.Text = stickStatus.ToString();
                lblProfileStep_Manual.Text = profileStep.ToString();
                lblRunTime_Manual.Text = runTimeSec.ToString();
                lblPVSUSTempP_Manual.Text = pvSusTempP_C.ToString();
                lblPVSUSTempA_Manual.Text = pvSusTempA_C.ToString();
                lblHeaterCurrent_Manual.Text = heaterCurrent_mA.ToString();
                lblHeaterVolt_Manual.Text = heaterVolt_mV.ToString();
                lblPVPCBTemp_Manual.Text = pvPcbTemp_C.ToString();
                lblPVDuty_Manual.Text = pvDuty_pct.ToString();
                lblBATTVolt_Manual.Text = battVolt_mV.ToString();
                lblBATTTemp_Manual.Text = battTemp_C.ToString();
                lblBATTSoc_Manual.Text = battSoC_pct.ToString();
            }

            return;
        }

        private async void btnCommandSend_Click(object sender, EventArgs e)
        {
            if (DevicePort == null || !DevicePort.IsOpen)
            {
                LogToUI("Device", "디바이스 포트가 열려있지 않습니다.", 0);
                return;
            }

            var cmd = cboxCommandList.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(cmd))
            {
                LogToUI("Device", "명령이 선택되지 않았습니다.", 0);
                return;
            }

            

            switch (cmd)
            {
                case "Charge OFF":
                    {
                        await Charge_Off_Manual();
                        break;
                    }

                case "Set Init R Init Temp":
                    {
                        await Set_Init_R_Init_Temp_Manual();
                        break;
                    }


                case "Set Calibration Profile":
                    {
                        await Set_Calibration_Profile_Manual();
                        break;
                    }

                case "Check Profile":
                    {
                        await Check_Profile_Manual();
                        break;
                    }
                case "Set TR Mode Off":
                    {
                        await Set_TR_Mode_Off_Manual();
                        break;
                    }
                case "ABC Value Reset": //절차 확인 필요 ==> 응답이 lo-hi가 바뀌어서 옴
                    {
                        await ABC_Value_Reset_Manual();
                        break;
                    }

                case "Select Use Profile":
                    {
                        await Select_Use_Profile_Manual();
                        break;
                    }

                case "Heating Start":
                    {
                        await Heating_Start_Manual();
                        break;
                    }

                case "Heating Stop":
                    {
                        await Heating_Stop_Manual();
                        break;
                    }
                default:
                    {
                        LogToUI("Device", $"알 수 없는 명령 : {cmd}", 0);
                        break;
                    }
            }
        }

        private async void btnGetDeviceData_Click(object sender, EventArgs e)
        {
            if (DevicePort == null || !DevicePort.IsOpen)
            {
                MessageBox.Show("디바이스 포트가 연결되어있지 않습니다.");
                return;
            }
            await Get_Device_Data_Manual();
        }

        private void btnCalTimeGet_Click(object sender, EventArgs e)
        {
            long point3 = Settings.Instance.Cal_Point_3_Time / 1000;
            long point2 = Settings.Instance.Cal_Point_2_Time / 1000;
            long point1 = Settings.Instance.Cal_Point_1_Time / 1000;
            tboxCalPoint3Time.Text = point3.ToString();
            tboxCalPoint2Time.Text = point2.ToString();
            tboxCalPoint1Time.Text = point1.ToString();
        }

        private void btnCalTimeSet_Click(object sender, EventArgs e)
        {
            long input_point3;
            long input_point2;
            long input_point1;

            if (string.IsNullOrEmpty(tboxCalPoint3Time.Text))
                input_point3 = 0;
            else
                input_point3 = long.Parse(tboxCalPoint3Time.Text) * 1000;

            if (string.IsNullOrEmpty(tboxCalPoint2Time.Text))
                input_point2 = 0;
            else
                input_point2 = long.Parse(tboxCalPoint2Time.Text) * 1000;

            if (string.IsNullOrEmpty(tboxCalPoint1Time.Text))
                input_point1 = 0;
            else
                input_point1 = long.Parse(tboxCalPoint1Time.Text) * 1000;
            
            Settings.Instance.Cal_Point_3_Time = input_point3;
            Settings.Instance.Cal_Point_2_Time = input_point2;
            Settings.Instance.Cal_Point_1_Time = input_point1;

            Settings.Instance.Save();
        }

        #endregion

        #region Maurer Manual
        private async void btnEmissivitySet_Click(object sender, EventArgs e)
        {
            try
            {
                if (MeasPort == null || !MeasPort.IsOpen)
                {
                    LogToUI("Maurer", "온도계가 연결되어있지 않습니다.", 0);
                    return;
                }

                // 예: 텍스트박스에 "80.0"℃ 입력 → tenthC=800
                double valC;
                if (!double.TryParse(tboxEmissivitySet.Text, out valC))
                {
                    LogToUI("Maurer", "보정값 입력 오류", 0);
                    return;
                }

                int tenthC = (int)(valC * 10);
                string resp = await MeasPort.Ce_SetAsync(tenthC, Meas_ReadTimeOut_value);

                if (resp == null)
                {
                    LogToUI("Maurer", "Calib Set : NULL", 0);
                    return;
                }

                LogToUI("Maurer", $"Calib Set {valC:F1}℃ → 응답:{resp}", 1);
                lblEmissivitySet.Text = $"{valC:F1}℃";
            }
            catch (Exception ex)
            {
                LogToUI("Maurer", $"Calib Set 예외: {ex.Message}", 0);
            }
        }

        private async void btnEmissivityLockSet_Click(object sender, EventArgs e)
        {
            try
            {
                if (MeasPort == null || !MeasPort.IsOpen)
                {
                    LogToUI("Maurer", "온도계가 연결되어있지 않습니다.", 0);
                    return;
                }

                // 체크박스나 토글버튼 값에서 설정 (예: chkLock.Checked)
                bool toLock = chkLock.Checked;

                string resp = await MeasPort.Lk_SetAsync(toLock, Meas_ReadTimeOut_value);

                if (resp == null)
                {
                    LogToUI("Maurer", "Lock Set : NULL", 0);
                    return;
                }

                LogToUI("Maurer", $"Lock Set → {(toLock ? "LOCKED" : "ACTIVE")} (응답:{resp})", 1);
                lblEmissivityLockSet.Text = $"{(toLock ? "→LOCKED" : "→ACTIVE")}";
            }
            catch (Exception ex)
            {
                LogToUI("Maurer", $"Lock Set 예외: {ex.Message}", 0);
            }
        }

        private async void btnEmissivityLockGet_Click(object sender, EventArgs e)
        {
            try
            {
                if (MeasPort == null || !MeasPort.IsOpen)
                {
                    LogToUI("Maurer", "온도계가 연결되어있지 않습니다.", 0);
                    return;
                }

                bool? locked = await MeasPort.Lk_GetAsync(Meas_ReadTimeOut_value);

                if (!locked.HasValue)
                {
                    lblEmissivityLockGet.Text = "NULL";
                    LogToUI("Maurer", "Lock State : NULL", 0);
                    return;
                }

                lblEmissivityLockGet.Text = locked.Value ? "LOCKED" : "ACTIVE";
                LogToUI("Maurer", $"Lock State : {lblEmissivityLockGet.Text}", 1);
            }
            catch (TimeoutException)
            {
                LogToUI("Maurer", "Lock State : TIMEOUT", 0);
            }
            catch (Exception ex)
            {
                LogToUI("Maurer", $"예외: {ex.Message}", 0);
            }
        }

       

        private async void btnNowEmissivity_Click(object sender, EventArgs e)
        {
            try
            {
                if (MeasPort == null || !MeasPort.IsOpen)
                {
                    LogToUI("Maurer", "온도계가 연결되어있지 않습니다.", 0);
                    return;
                }

                double? emiss = await MeasPort.Em_ReadAsync(Meas_ReadTimeOut_value);

                if (!emiss.HasValue)
                {
                    lblNowEmissivity.Text = "NULL";
                    LogToUI("Maurer", "Emissivity : NULL", 0);
                    return;
                }

                lblNowEmissivity.Text = $"{emiss.Value:F3}";
                LogToUI("Maurer", $"Emissivity : {emiss.Value:F3}", 1);
            }
            catch (TimeoutException)
            {
                LogToUI("Maurer", "Emissivity : TIMEOUT", 0);
            }
            catch (Exception ex)
            {
                LogToUI("Maurer", $"예외: {ex.Message}", 0);
            }
        }

        private async void btnNowTemp_Click(object sender, EventArgs e)
        {
            try
            {
                if (MeasPort == null || !MeasPort.IsOpen)
                {
                    LogToUI("Maurer", "온도계가 연결되어있지 않습니다.", 0);
                    return;
                }

                double? temp = await MeasPort.Ms_ReadTemp(Meas_ReadTimeOut_value);

                if (!temp.HasValue)
                {
                    lblMaurerTempManual.Text = "NULL";
                    LogToUI("Maurer", "Maurer Temp : NULL", 0);
                    return; // ← 여기서 끝내야 아래서 덮어쓰지 않음
                }

                lblMaurerTempManual.Text = $"{temp.Value:F1} ℃";
                LogToUI("Maurer", $"Maurer Temp : {temp.Value:F1} ℃", 1);
            }
            catch (TimeoutException)
            {
                Console.WriteLine("CommandSend ReadTimeOut Exception!");
                LogToUI("Maurer", "Maurer Temp : TIMEOUT", 0);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Command Send 중 예외 발생! {ex.Message}");
                LogToUI("Maurer", $"예외: {ex.Message}", 0);
            }
        }




        #endregion

        #region Calibration
        private async void PollTimer90_Tick(object sender, EventArgs e)
        {
            if (_runCts?.IsCancellationRequested == true) return;

            // 1) 게이트 잡고 읽기/표시만 빠르게
            if (!await _ioGate.WaitAsync(0)) return;
            List<ScheduledAction> due = null;

            try
            {
                var ct = _runCts?.Token ?? CancellationToken.None;
                bool devOn = DevicePort?.IsOpen == true;
                bool measOn = MeasPort?.IsOpen == true;

                double elapsedSec = _sw.ElapsedMilliseconds / 1000.0;
                var dev = DEV_EMPTY;
                double? maurer = null;
                
                if (elapsedSec < 91) //90초 이후엔 온도받아오는거 안한다. (프로파일이 90초까지임)
                {
                    var devTask = devOn ? ReadDeviceSnapshotAsync(90, ct) : Task.FromResult(DEV_EMPTY);

                    var measTask = measOn ? MeasPort.Ms_ReadTemp(90) : Task.FromResult<double?>(null);

                    await Task.WhenAll(devTask, measTask);

                    dev = devTask.Status == TaskStatus.RanToCompletion ? devTask.Result : DEV_EMPTY;
                    maurer = measTask.Status == TaskStatus.RanToCompletion ? measTask.Result : (double?)null;

                    lock (_latestLock)
                    {
                        if (dev.ok)
                        {
                            _lastTempP = dev.tp; _lastTempA = dev.ta;
                            _lastRunTimeSec = dev.runSec; _lastProfileStep = dev.step;
                        }
                        if (maurer.HasValue) _lastMaurer = maurer.Value;
                    }


                    UpdateUiAndGraph();
                }
                
                
                

                //==================Point 3 // X3 Y3 산출=====================================================
                if (_p3Active && dev.ok && maurer.HasValue)
                {
                    //double elapsedSec = _sw.ElapsedMilliseconds / 1000.0;
                    

                    if (elapsedSec >= P3_SAMPLE_START && elapsedSec < P3_SAMPLE_END)
                    {
                        _p3TempP.Add(dev.tp);
                        _p3Thermo.Add(maurer.Value);
                    }

                    if (!_p3Averaged && elapsedSec >= P3_SAMPLE_END)
                    {
                        X3 = (float)TrimmedMean(_p3TempP);
                        Y3 = (float)TrimmedMean(_p3Thermo);
                        _p3Averaged = true;
                        _p3Active = false;


                        Console.WriteLine($"X3 = {X3} Y3 = {Y3}");
                        LogToUI("Cal", $"[Point 3] Samples Count Temp_P : {_p3TempP.Count} Maurer Temp : {_p3Thermo.Count}", 1);
                        LogToUI("Cal", $"[Point 3] X3 = {X3} Y3 = {Y3}", 1);
                    }
                }
                //=================================================================================================

                //==================Point 2 // X2 Y2 산출=====================================================
                if (_p2Active && dev.ok && maurer.HasValue)
                {
                    //double elapsedSec = _sw.ElapsedMilliseconds / 1000.0;


                    if (elapsedSec >= P2_SAMPLE_START && elapsedSec < P2_SAMPLE_END)
                    {
                        _p2TempP.Add(dev.tp);
                        _p2Thermo.Add(maurer.Value);
                    }

                    if (!_p2Averaged && elapsedSec >= P2_SAMPLE_END)
                    {
                        X2 = (float)TrimmedMean(_p2TempP);
                        Y2 = (float)TrimmedMean(_p2Thermo);
                        _p2Averaged = true;
                        _p2Active = false;

                        Console.WriteLine($"X2 = {X2} Y2 = {Y2}");
                        LogToUI("Cal", $"[Point 2] Samples Count Temp_P : {_p2TempP.Count} Maurer Temp : {_p2Thermo.Count}", 1);
                        LogToUI("Cal", $"[Point 2] X2 = {X2} Y2 = {Y2}", 1);
                    }
                }
                //=================================================================================================

                //==================Point 1 // X1 Y1 산출=====================================================
                if (_p1Active && dev.ok && maurer.HasValue)
                {
                    //double elapsedSec = _sw.ElapsedMilliseconds / 1000.0;


                    if (elapsedSec >= P1_SAMPLE_START && elapsedSec < P1_SAMPLE_END)
                    {
                        _p1TempP.Add(dev.tp);
                        _p1Thermo.Add(maurer.Value);
                    }

                    if (!_p1Averaged && elapsedSec >= P1_SAMPLE_END)
                    {
                        X1 = (float)TrimmedMean(_p1TempP);
                        Y1 = (float)TrimmedMean(_p1Thermo);
                        _p1Averaged = true;
                        _p1Active = false;

                        Console.WriteLine($"X1 : {X1} Y1 : {Y1}");
                        LogToUI("Cal", $"[Point 1] Samples Count Temp_P : {_p1TempP.Count} Maurer Temp : {_p1Thermo.Count}", 1);
                        LogToUI("Cal", $"[Point 1] X1 = {X1} Y1 = {Y1}", 1);
                    }
                }
                //=================================================================================================

                // 2) 지금 시각 기준 실행해야 할 스케줄 항목만 복사 (게이트 안에서 큐만 긁어옴)
                long nowMs = _sw.ElapsedMilliseconds;
                due = _schedule.Where(s => !s.Done && nowMs >= s.DueMs).ToList();
                //foreach (var s in due) s.Done = true; // 중복 실행 방지 마킹
            }
            catch (Exception ex)
            {
                LogToUI("Device", $"Poll error: {ex.Message}", 0);
            }
            finally
            {
                _ioGate.Release(); // <= 폴링/표시 끝났으니 바로 해제
            }

            // 3) 스케줄은 게이트 밖에서, 백그라운드 태스크로 실행
            if (due != null && due.Count > 0)
            {
                foreach (var s in due)
                {
                    _ = Task.Run(async () =>
                    {
                        try
                        {
                            // 액션 내부에서 포트 I/O가 필요하면 s.Action 안에서
                            // await _ioGate.WaitAsync(); try { ... } finally { _ioGate.Release(); }
                            // 패턴으로 잠깐씩만 잡도록 설계
                            await s.Action(_runCts?.Token ?? CancellationToken.None);
                        }
                        catch (Exception ex)
                        {
                            LogToUI("Device", $"Scheduled action error: {ex.Message}", 0);
                        }
                        finally
                        {
                            s.Done = true;
                        }
                    });
                }
            }
            
            if (_schedule.Count > 0 && _schedule.All(s => s.Done))
            {
                Console.WriteLine("모든 작업 종료");
                _ = Task.Run(async() => await StopMain());
                //await StopMain();
            }//다 끝나면 종료
                
        }

        private void UpdateUiAndGraph()
        {
            double? tp, ta, mt;
            double run;
            ushort step;

            lock (_latestLock)
            {
                tp = _lastTempP;
                ta = _lastTempA;
                mt = _lastMaurer;
                run = _lastRunTimeSec;
                step = _lastProfileStep;
            }

            void SetIfExists(System.Windows.Forms.Label lbl, string text)
            {
                if (lbl != null) lbl.Text = text;
            }

            // 폼에 실제 있는 컨트롤 이름에 맞춰 호출
            SetIfExists(lblRunTime, run.ToString("F1"));
            SetIfExists(lblProfileStep, step.ToString());
            SetIfExists(lblTemp_P, tp?.ToString("F2") ?? "");
            SetIfExists(lblTemp_A, ta?.ToString("F2") ?? "");
            SetIfExists(lblMaurerTemp, mt?.ToString("F2") ?? "");

            // 그래프 X축: Time(x100ms) → ms/100
            double x = _sw.ElapsedMilliseconds / 100.0;

            bool any = false;
            if (tp.HasValue) { _listP.Add(x, tp.Value); any = true; }
            if (ta.HasValue) { _listA.Add(x, ta.Value); any = true; }
            if (mt.HasValue) { _listExt.Add(x, mt.Value); any = true; }

            if (any)
            {
                // 슬라이딩 윈도우 유지 (0~3000)
                var p = TempGraph.GraphPane;
                if (x > X_WINDOW)
                {
                    p.XAxis.Scale.Min = x - X_WINDOW;
                    p.XAxis.Scale.Max = x;
                }
                TempGraph.AxisChange();
                TempGraph.Invalidate();
            }
        }

        private async Task<(bool ok, double tp, double ta, double runSec, ushort step)> ReadDeviceSnapshotAsync(int timeoutMs, CancellationToken ct)
        {
            try
            {
                if (DevicePort == null || !DevicePort.IsOpen) return (false, 0, 0, 0, 0);

                byte[] tx = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.GET_DEVICE_DATA).GetPacket();
                byte[] rx = await DevicePort.SendAndReceivePacketAsync(tx, timeoutMs);
                if (rx == null) return (false, 0, 0, 0, 0);

                if (!TryParseAsciiToBytes(rx, out var frame)) frame = rx;
                if (frame.Length < 3) return (false, 0, 0, 0, 0);

                int byteCount = frame[2];
                int PayloadOffset = 3;
                if (byteCount != 26 || frame.Length < PayloadOffset + byteCount) return (false, 0, 0, 0, 0);

                ushort U16(int reg) { int i = PayloadOffset + reg * 2; return (ushort)(frame[i] | (frame[i + 1] << 8)); }

                var status = U16(0); // not used here
                var stickStatus = U16(1); // not used here
                var profileStep = U16(2);
                var runTimeSec = U16(3) / 10.0;

                var tp = U16(4) / 10.0 - 100.0; // P
                var ta = U16(5) / 10.0 - 100.0; // A

                return (true, tp, ta, runTimeSec, profileStep);
            }
            catch { return (false, 0, 0, 0, 0); }
        }

        private async void btnCalStart_Click(object sender, EventArgs e)
        {
            if (DevicePort == null || !DevicePort.IsOpen)
            {
                MessageBox.Show("디바이스 포트를 연결해주세요");
                return;
            }

            if (MeasPort == null || !MeasPort.IsOpen)
            {
                MessageBox.Show("온도계 포트를 연결해주세요");
                return;
            }

            if(Settings.Instance.Use_Write_Log)
                CsvManager.CsvSave(Environment.NewLine + "==================== < Calibration Start > ====================" + Environment.NewLine);

            DevicePort.LogCommToUI = null;
            MeasPort.LogCommToUI = null;

            bool isPreSettingOk = false;
            isPreSettingOk = await Pre_Cal_Sequence_Task();
            if (!isPreSettingOk) return;
            _taskKind = TaskKind.Calibration;

            await Heating_Start();

            StartMain();
        }

        private async void btnCalStop_Click(object sender, EventArgs e)
        {
            deliberate_stop = true;
            await StopMain();
            deliberate_stop = false;
        }

        private async Task StopMain()
        {
            _pollTimer90.Stop();
            await Heating_Stop();
            if (!deliberate_stop) await SendToDbData(_taskKind);
            _runCts?.Cancel();
            _runCts?.Dispose();
            _runCts = null;           // ← 추가
            _schedule.Clear();        // 선택: 스케줄도 초기화
            DevicePort.LogCommToUI = LogCommToUI;
            MeasPort.LogCommToUI = LogCommToUI;
            if (Settings.Instance.Use_Write_Log)
                CsvManager.CsvSave(Environment.NewLine + "==================== < Calibration End > ====================" + Environment.NewLine);
        }
        private void EnqueueAt(long dueMs, Func<CancellationToken, Task> action)
        {
            _schedule.Add(new ScheduledAction { DueMs = dueMs, Action = action, Done = false });
            _schedule.Sort((a, b) => a.DueMs.CompareTo(b.DueMs));
        }

        private async Task WithIo(Func<Task> body, CancellationToken ct = default)
        {
            if (ct.IsCancellationRequested) return;

            await _ioGate.WaitAsync(ct);
            try { await body(); }
            finally { _ioGate.Release(); }
        }
        private void StartMain()
        {
            if (_runCts != null) return;              // 이미 동작 중이면 무시
            _runCts = new CancellationTokenSource();

            // 그래프/시간 초기화
            _listA.Clear(); _listP.Clear(); _listExt.Clear();
            _sw.Restart();                             // 기준 시간 0으로
            TempGraph.AxisChange(); TempGraph.Invalidate();

            //결과 플래그 초기화
            Cal_Result_flag = 1;

            // 스케줄 구성
            _schedule.Clear();
            ResetPoint3();
            ResetPoint2();
            ResetPoint1();

            EnqueueAt(Settings.Instance.Cal_Point_3_Time, ct => WithIo(() => WritePoint3DiffScheduledAsync(ct), ct));
            EnqueueAt(Settings.Instance.Cal_Point_2_Time, ct => WithIo(() => WritePoint2DiffScheduledAsync(ct), ct));
            EnqueueAt(Settings.Instance.Cal_Point_1_Time, ct => WithIo(() => WritePoint1DiffScheduledAsync(ct), ct));
            

            EnqueueAt(ABC_Write_Time, ct => WithIo(() => ABC_Value_Write(ct), ct));

            EnqueueAt(Set_Verification_Profile_Time, ct => WithIo(() => Set_Verification_Profile(ct),ct));
            EnqueueAt(Set_TR_Mode_Time, ct => WithIo(() => Set_TR_Mode(ct), ct)); //coconut 모델 적용
            EnqueueAt(Set_Calibration_Flag_TIme, ct => WithIo(() => Set_Calibration_Flag(ct), ct));
            EnqueueAt(Check_Verification_Profile_Time, ct => WithIo(() => Check_Verification_Profile(ct), ct));
            EnqueueAt(Flash_Update_Time, ct => WithIo(() => Flash_Update(ct), ct));
            //EnqueueAt(100_000, ct => WithIo(() => Heating_Stop(), ct));
            // 90ms 폴링 시작
            _pollTimer90.Interval = 90;
            _pollTimer90.Start();

            //void Enq(long ms, Func<CancellationToken, Task> act) =>
            //    _schedule.Add(new ScheduledAction { DueMs = ms, Action = act, Done = false });
        }

        #endregion 

        #region Calibration Task Function
        private async Task<bool> Pre_Cal_Sequence_Task()
        {
            bool isOk = false;
            try
            {
                isOk = await Charge_Off(); 
                if (!isOk)
                {
                    LogToUI("Cal", "=============================< Charge Off 실패 >=================================", 0);
                    LogToUI("Cal", "=============================< Calibration 사전작업 실패 >=================================", 0);
                    return false;
                }

                isOk = await Set_Init_R_Init_Temp();
                if (!isOk)
                {
                    LogToUI("Cal", "=============================< Set Init R Init Temp 실패 >=================================", 0);
                    LogToUI("Cal", "=============================< Calibration 사전작업 실패 >=================================", 0);
                    return false;
                }
                isOk = await Set_Calibration_Profile();
                if (!isOk)
                {
                    LogToUI("Cal", "=============================< Set Calibration Profile 실패 >=================================", 0);
                    LogToUI("Cal", "=============================< Calibration 사전작업 실패 >=================================", 0);
                    return false;
                }
                isOk = await Check_Profile();
                if (!isOk)
                {
                    LogToUI("Cal", "=============================< Check Profile 실패 >=================================", 0);
                    LogToUI("Cal", "=============================< Calibration 사전작업 실패 >=================================", 0);
                    return false;
                }
                isOk = await Set_TR_Mode_Off();
                if (!isOk)
                {
                    LogToUI("Cal", "=============================< Set TR Mode Off 실패 >=================================", 0);
                    LogToUI("Cal", "=============================< Calibration 사전작업 실패 >=================================", 0);
                    return false;
                }
                isOk = await ABC_Value_Reset();
                if (!isOk)
                {
                    LogToUI("Cal", "=============================< ABC Value Reset 실패 >=================================", 0);
                    LogToUI("Cal", "=============================< Calibration 사전작업 실패 >=================================", 0);
                    return false;
                }
                isOk = await Select_Use_Profile();
                if (!isOk)
                {
                    LogToUI("Cal", "=============================< Select Use Profile 실패 >=================================", 0);
                    LogToUI("Cal", "=============================< Calibration 사전작업 실패 >=================================", 0);
                    return false;
                }
                isOk = await Charge_Off();
                if (!isOk)
                {
                    LogToUI("Cal", "=============================< Charge Off 실패(2) >=================================", 0);
                    LogToUI("Cal", "=============================< Calibration 사전작업 실패 >=================================", 0);
                    return false;
                }

                
                LogToUI("Cal", "=============================< Calibration 사전작업 완료 >=================================", 1);
                return isOk;
            }
            catch (Exception ex)
            {
                LogToUI("Cal", $"{ex.Message}", 0);
                LogToUI("Cal", "=============================< Calibration 사전작업 실패 >=================================", 0);
                isOk = false;
                return isOk;
            } 
        }
        private async Task<bool> Charge_Off()
        {
            int timeout = Device_ReadTimeOut_Value;
            // 1) 패킷 생성
            byte[] tx = new CDCProtocol(Variable.SLAVE, Variable.WRITE, Variable.CHARGE_OFF).GetPacket();

            // 2) 송수신
            byte[] rx = await DevicePort.SendAndReceivePacketAsync(tx, timeout);
            if (rx == null)
            {
                LogToUI("Cal", "응답 없음", 0);
                return false;
            }

            // 3) 응답 검증 및 로그
            string why;
            if (IsEchoAck(tx, rx, out why))
            {
                LogToUI("Cal", "Charge OFF 적용 완료", 1);
                return true;
            }
                
            else
            {
                LogToUI("Cal", "Charge OFF 적용 실패", 0);
                return false;
            }
 
        }
        private async Task<bool> Set_Init_R_Init_Temp()
        {
            int timeout = Device_ReadTimeOut_Value;
            // 1) 패킷 생성
            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.MULTI_WRITE, Variable.SET_INIT_R_INIT_TEMP).GetPacket();

            // 2) 송수신
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            if (rx1 == null || !IsWriteMultiAck(rx1, Variable.SLAVE, 0x00AF, 0x0002))
            {
                LogToUI("Cal", "Set init R init temp 실패(멀티 쓰기 ACK 불일치)", 0);
                return false;
            }

            byte[] tx2 = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.CHECK_INIT_R_INIT_TEMP).GetPacket();
            byte[] rx2 = await DevicePort.SendAndReceivePacketAsync(tx2, timeout);
            if (rx2 == null)
            {
                LogToUI("Cal", "Check init R init Temp 응답 없음", 0);
                return false;
            }

            byte[] frame;
            if (!TryParseAsciiToBytes(rx2, out frame))
                frame = rx2;

            if (frame.Length < 3)
            {
                LogToUI("Cal", "Check init R init Temp 실패 (응답 프레임 짧음)", 0);
                return false;
            }

            int byteCount = frame[2];
            int payloadOffset = 3;

            if (frame.Length < payloadOffset + byteCount)
            {
                LogToUI("Cal", "Check init R init Temp 실패 (응답 ByteCount 불일치)", 0);
                return false;
            }

            if (byteCount != 4) //레지스터 2개 == 4바이트
            {
                LogToUI("Cal", $"Check Profile 실패 (예상 ByteCount=4, 수신={byteCount})", 0);
                return false;
            }

            byte[] Byte_Heater_Init_Resistance = { frame[3], frame[4] };
            byte[] Byte_Heater_Init_Temperature = { frame[5], frame[6] };
            short Heater_Init_Resistance = BitConverter.ToInt16(Byte_Heater_Init_Resistance, 0);
            short Heater_Init_Temperature = BitConverter.ToInt16(Byte_Heater_Init_Temperature, 0);
            //byte[] Byte_expected_Resistance = { 0x00, 0x00 };
            byte[] Byte_expected_Temperature = { Variable.SET_INIT_R_INIT_TEMP[7], Variable.SET_INIT_R_INIT_TEMP[8] };
            //Array.Reverse(Byte_expected_Resistance);
            Array.Reverse(Byte_expected_Temperature);
            //short expected_Heater_Init_Resistance = BitConverter.ToInt16(Byte_expected_Resistance, 0);
            short expected_Heater_Init_Temperature = BitConverter.ToInt16(Byte_expected_Temperature, 0);

            short Resistance_Min = 0;
            short Resistance_Max = 2000;

            bool isResisOk = Heater_Init_Resistance > Resistance_Min && Heater_Init_Resistance < Resistance_Max;
            bool isTempOk = Heater_Init_Temperature == expected_Heater_Init_Temperature;

            if (!isResisOk || !isTempOk)
            {
                LogToUI("Cal", "Set init R init Temp 실패", 0);
                LogToUI("Cal", $"Resistance ==> Set = {Heater_Init_Resistance} [{Resistance_Min} ~ {Resistance_Max}]", isResisOk ? 1 : 0);
                LogToUI("Cal", $"Temperature ==> Set = {Heater_Init_Temperature} [{expected_Heater_Init_Temperature}]", isTempOk ? 1 : 0);
                return false;
            }

            LogToUI("Cal", "Set init R init Temp 완료", 1);
            LogToUI("Cal", $"Set = {Heater_Init_Resistance} [{Resistance_Min} ~ {Resistance_Max}]", isResisOk ? 1 : 0);
            LogToUI("Cal", $"Set = {Heater_Init_Temperature} [{expected_Heater_Init_Temperature}]", isTempOk ? 1 : 0);

            return true;
        }
        private async Task<bool> Set_Calibration_Profile()
        {
            if (string.IsNullOrEmpty(Settings.Instance.Profile_File) || string.IsNullOrEmpty(Settings.Instance.Profile_File_Path))
            {
                MessageBox.Show("프로파일을 선택해주세요.");
                return false;
            }
            int timeout = Device_ReadTimeOut_Value;
            // 1) Settings.Instance 값으로 데이터 빌드 (온도12 + 시간12 → 총 24워드)
            byte[] data = BuildProfileData(true);
            byte[] payload = new byte[2 + 2 + 1 + data.Length];

            payload[0] = 0x00; payload[1] = 0x55;
            payload[2] = 0x00; payload[3] = 0x18;
            payload[4] = (byte)data.Length;

            Buffer.BlockCopy(data, 0, payload, 5, data.Length);

            // 2) 0x55 ~ 0x6C 연속 쓰기 (Write Multiple Registers)
            //    ↓ CDCProtocol 생성자는 네 프로젝트 규약에 맞게 수정해
            //    예시: (slave, func: WRITE_MULTI, startAddr, payload)
            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.MULTI_WRITE, payload).GetPacket();
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            if (rx1 == null || !IsWriteMultiAck(rx1, Variable.SLAVE, 0x0055, 0x0018))
            {
                LogToUI("Cal", "프로파일 전송 실패(멀티 쓰기 ACK 불일치)", 0);
                return false;
            }

            // 3) 0x6D에 0x0101 쓰기 (프로파일 인덱스/적용 트리거)
            //    예: (slave, func: WRITE, address, data2바이트)
            byte[] applyData = new byte[4] { 0x00, 0x6D, 0x01, 0x01 }; // 0x0101 (big-endian 전송이면 CDCProtocol에서 처리)
            byte[] tx2 = new CDCProtocol(Variable.SLAVE, Variable.WRITE, applyData).GetPacket();
            byte[] rx2 = await DevicePort.SendAndReceivePacketAsync(tx2, timeout);
            if (rx2 == null)
            {
                LogToUI("Cal", "프로파일 적용 응답 없음", 0);
                return false;
            }

            string why2;
            if (IsEchoAck(tx2, rx2, out why2))
            {
                LogToUI("Cal", "Set Calibration Profile 완료", 1);
                return true;
            }
                
            else
            {
                LogToUI("Cal", "Set Calibration Profile 실패", 0);
                return false;
            }
                

           
        }
        private async Task<bool> Check_Profile()
        {
            int timeout = Device_ReadTimeOut_Value;
            byte[] applyData = new byte[4] { 0x00, 0x6D, 0x00, 0x01 };
            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.WRITE, applyData).GetPacket();
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            string why1;
            if (rx1 == null || !IsEchoAck(tx1, rx1, out why1))
            {
                LogToUI("Cal", "Check Profile 실패", 0);
                return false;
            }

            byte[] tx2 = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.CHECK_PROFILE).GetPacket();
            byte[] rx2 = await DevicePort.SendAndReceivePacketAsync(tx2, timeout);
            if (rx2 == null)
            {
                LogToUI("Cal", "Check Profile 실패 (프로파일 Read 응답 없음)", 0);
                return false;
            }

            byte[] frame;
            if (!TryParseAsciiToBytes(rx2, out frame))
                frame = rx2;

            if (frame.Length < 3)
            {
                LogToUI("Cal", "Check Profile 실패 (응답 프레임 짧음)", 0);
                return false;
            }

            int byteCount = frame[2];
            int payloadOffset = 3;

            if (frame.Length < payloadOffset + byteCount)
            {
                LogToUI("Cal", "Check Profile 실패 (응답 ByteCount 불일치)", 0);
                return false;
            }

            if (byteCount != 48) // 24워드 * 2바이트
            {
                LogToUI("Cal", $"Check Profile 실패 (예상 ByteCount=48, 수신={byteCount})", 0);
                return false;
            }

            // 기대값 (big-endian → WORD)
            byte[] expected = BuildProfileData(true); // 48 bytes
            ushort[] expectedWords = new ushort[24];
            for (int i = 0; i < 24; i++)
            {
                expectedWords[i] = (ushort)((expected[i * 2] << 8) | expected[i * 2 + 1]);
            }

            // 응답값 (장치가 lo,hi 순서로 보냄 → 리틀엔디언 스왑)
            ushort[] receivedWords = new ushort[24];
            for (int i = 0; i < 24; i++)
            {
                int idx = payloadOffset + i * 2;
                receivedWords[i] = (ushort)(frame[idx] | (frame[idx + 1] << 8)); // lo | (hi<<8)
            }

            bool same = true;
            string tempLog = "온도 값 비교:\r\n";
            string timeLog = "시간 값 비교:\r\n";

            for (int i = 0; i < 12; i++)
            {
                ushort expT = expectedWords[i];        // 기대 온도 i+1
                ushort expTi = expectedWords[12 + i];   // 기대 시간 i+1

                ushort actT = receivedWords[i * 2];     // 읽은 온도 i+1 (짝수 인덱스)
                ushort actTi = receivedWords[i * 2 + 1]; // 읽은 시간 i+1 (홀수 인덱스)

                tempLog += string.Format("Point{0}_Temp: 설정={1}, 읽음={2}\r\n", i + 1, expT, actT);
                timeLog += string.Format("Point{0}_Time: 설정={1}, 읽음={2}\r\n", i + 1, expTi, actTi);

                if (expT != actT || expTi != actTi)
                    same = false;
            }

            LogToUI("Cal",
                (same ? "Check Profile OK\r\n" : "Check Profile 불일치\r\n") + tempLog + timeLog,
                same ? 1 : 0);

            return same;
        }
        private async Task<bool> Set_TR_Mode_Off()
        {
            int timeout = Device_ReadTimeOut_Value;
            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.MULTI_WRITE, Variable.SET_TR_MODE_OFF).GetPacket();
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            if (rx1 == null || !IsWriteMultiAck(rx1, Variable.SLAVE, 0x0055, 0x000C))
            {
                LogToUI("Cal", "SET TR MODE OFF 실패(멀티 쓰기 ACK 불일치)", 0);
                return false;
            }

            byte[] applyData = new byte[4] { 0x00, 0x6D, 0x01, 0x00 };
            byte[] tx2 = new CDCProtocol(Variable.SLAVE, Variable.WRITE, applyData).GetPacket();
            byte[] rx2 = await DevicePort.SendAndReceivePacketAsync(tx2, timeout);
            if (rx2 == null)
            {
                LogToUI("Cal", "데이터 쓰기 응답 없음", 0);
                return false;
            }

            string why2;
            if (!IsEchoAck(tx2, rx2, out why2))
            {
                LogToUI("Cal", "Set TR Mode Off 실패", 0);
                return false;
            }
                


            applyData = new byte[4] { 0x00, 0x6D, 0x00, 0x00 }; //0x6D에 0x0000 write
            byte[] tx3 = new CDCProtocol(Variable.SLAVE, Variable.WRITE, applyData).GetPacket();
            byte[] rx3 = await DevicePort.SendAndReceivePacketAsync(tx3, timeout);
            string why3;
            if (rx3 == null || !IsEchoAck(tx3, rx3, out why3))
            {
                LogToUI("Cal", "Check Profile 실패", 0);
                return false;
            }

            byte[] tx4 = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.CHECK_TR_MODE).GetPacket();
            byte[] rx4 = await DevicePort.SendAndReceivePacketAsync(tx4, timeout);
            if (rx4 == null)
            {
                LogToUI("Cal", "Check TR Mode 실패 (Read 응답 없음)", 0);
                return false;
            }

            byte[] frame;
            if (!TryParseAsciiToBytes(rx4, out frame)) frame = rx4;

            if (frame.Length < 3)
            {
                LogToUI("Cal", "Check TR Mode 실패 (응답 프레임 짧음)", 0);
                return false;
            }

            int byteCount = frame[2];
            int payloadOffset = 3;
            if (byteCount != 24 || frame.Length < payloadOffset + byteCount)
            {
                LogToUI("Cal", $"Check TR Mode 실패 (ByteCount=24 예상, 수신={byteCount})", 0);
                return false;
            }

            ushort[] temps = new ushort[12];
            for (int i = 0; i < 12; i++)
            {
                int idx = payloadOffset + i * 2;
                temps[i] = (ushort)(frame[idx] | (frame[idx + 1] << 8));
            }


            ushort[] expTemps = new ushort[12]; expTemps[3] = 30;

            bool same = true;
            var sb = new StringBuilder();
            sb.AppendLine("TR Mode Off 값 비교");
            for (int i = 0; i < 12; i++)
            {
                sb.AppendFormat("Point{0}_Temp: 설정={1}, 읽음={2}\r\n", i + 1, expTemps[i], temps[i]);
                if (expTemps[i] != temps[i]) same = false;
            }

            LogToUI("Cal", (same ? "Check TR Mode OK\n" : "Check TR Mode 불일치\n") + sb.ToString(), same ? 1 : 0);

            return same;
        }
        private async Task<bool> ABC_Value_Reset()
        {
            int timeout = Device_ReadTimeOut_Value;
            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.MULTI_WRITE, Variable.ABC_RESET_WRITE).GetPacket();
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            if (rx1 == null || !IsWriteMultiAck(rx1, Variable.SLAVE, 0x000E, 0x0006))
            {
                LogToUI("Cal", "ABC Value Reset 실패 (멀티쓰기 ACK 불일치)", 0);
                return false;
            }

            byte[] tx2 = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.ABC_RESET_READ).GetPacket();
            byte[] rx2 = await DevicePort.SendAndReceivePacketAsync(tx2, timeout);
            if (rx2 == null)
            {
                LogToUI("Cal", "ABC Value Reset 실패 (Read 응답 없음)", 0);
                return false;
            }

            byte[] frame;
            if (!TryParseAsciiToBytes(rx2, out frame)) frame = rx2;


            if (frame.Length < 3)
            {
                LogToUI("Cal", "ABC Value Reset 실패 (응답 프레임 짧음)", 0);
                return false;
            }

            int payloadOffset = 3;
            int byteCount = frame[2];
            if (byteCount != 12 || frame.Length < payloadOffset + byteCount)
            {
                LogToUI("Cal", $"ABC Value Reset 실패 (ByteCount=12 예상, 수신={byteCount})", 0);
                return false;
            }

            byte[] Byte_A = { frame[4], frame[3], frame[6], frame[5] };
            byte[] Byte_B = { frame[8], frame[7], frame[10], frame[9] };
            byte[] Byte_C = { frame[12], frame[11], frame[14], frame[13] };
            float A = BitConverter.ToSingle(Byte_A, 0);
            float B = BitConverter.ToSingle(Byte_B, 0);
            float C = BitConverter.ToSingle(Byte_C, 0);


            byte[] Byte_expected_A = { Variable.ABC_RESET_WRITE[5], Variable.ABC_RESET_WRITE[6], Variable.ABC_RESET_WRITE[7], Variable.ABC_RESET_WRITE[8] };
            byte[] Byte_expected_B = { Variable.ABC_RESET_WRITE[9], Variable.ABC_RESET_WRITE[10], Variable.ABC_RESET_WRITE[11], Variable.ABC_RESET_WRITE[12] };
            byte[] Byte_expected_C = { Variable.ABC_RESET_WRITE[13], Variable.ABC_RESET_WRITE[14], Variable.ABC_RESET_WRITE[15], Variable.ABC_RESET_WRITE[16] };

            float expected_A = BitConverter.ToSingle(Byte_expected_A, 0);
            float expected_B = BitConverter.ToSingle(Byte_expected_B, 0);
            float expected_C = BitConverter.ToSingle(Byte_expected_C, 0);

            Console.WriteLine($"return : {BitConverter.ToString(Byte_A)} expected : {BitConverter.ToString(Byte_expected_A)}");
            Console.WriteLine($"return : {BitConverter.ToString(Byte_B)} expected : {BitConverter.ToString(Byte_expected_B)}");
            Console.WriteLine($"return : {BitConverter.ToString(Byte_C)} expected : {BitConverter.ToString(Byte_expected_C)}");

            bool IsAOk = A == expected_A;
            bool IsBOk = B == expected_B;
            bool IsCOk = C == expected_C;

            if (!IsAOk || !IsBOk || !IsCOk)
            {
                LogToUI("Cal", "ABC Value Reset Fail", 0);
                LogToUI("Cal", $"A : 응답 = {A} 기대 = {expected_A} B : 응답 = {B} 기대 = {expected_B} C : 응답 = {C} 기대 = {expected_C}", 0);
                return false;
            }

            LogToUI("Cal", "ABC Value Reset Success", 1);
            LogToUI("Cal", $"A : 응답 = {A} 기대 = {expected_A} B : 응답 = {B} 기대 = {expected_B} C : 응답 = {C} 기대 = {expected_C}", 1);

            return true;
        }
        private async Task<bool> Select_Use_Profile()
        {
            int timeout = Device_ReadTimeOut_Value;
            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.WRITE, Variable.SELECT_USE_PROFILE_WRITE).GetPacket();
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            string why1;
            if (rx1 == null || !IsEchoAck(tx1, rx1, out why1))
            {
                LogToUI("Cal", "Select Use Profile Write 실패", 0);
                return false;
            }

            byte[] tx2 = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.SELECT_USE_PROFILE_READ).GetPacket();
            byte[] rx2 = await DevicePort.SendAndReceivePacketAsync(tx2, timeout);

            if (rx2 == null)
            {
                LogToUI("Cal", "Select Use Profile Read 실패", 0);
                return false;
            }

            byte[] frame;
            if (!TryParseAsciiToBytes(rx2, out frame)) frame = rx2;

            if (frame.Length < 3)
            {
                LogToUI("Cal", "Select Use Profile Read 실패 (응답 프레임 짧음)", 0);
                return false;
            }

            int payloadOffset = 3;
            int byteCount = frame[2];
            if (byteCount != 2 || frame.Length < payloadOffset + byteCount)
            {
                LogToUI("Cal", $"Select Use Profile Read 실패 (ByteCount=2 예상, 수신={byteCount})", 0);
                return false;
            }
            //write 빅엔디안 read 리틀엔디안?
            byte[] Byte_Profile_Select_Index = { frame[payloadOffset], frame[payloadOffset + 1] };
            ushort Profile_Select_Index = BitConverter.ToUInt16(Byte_Profile_Select_Index, 0); //받는거 리틀엔디안
            byte[] Byte_expected_Result = { Variable.SELECT_USE_PROFILE_WRITE[2], Variable.SELECT_USE_PROFILE_WRITE[3] };
            Array.Reverse(Byte_expected_Result); //주는거 빅엔디안
            ushort expected_Result = BitConverter.ToUInt16(Byte_expected_Result, 0);


            bool IsOk = Profile_Select_Index == expected_Result;

            if (!IsOk)
            {
                LogToUI("Cal", $"Select Use Profile 실패", 0);
                LogToUI("Cal", $"적용 : {Profile_Select_Index} [{expected_Result}]", 0);
                return false;
            }

            LogToUI("Cal", $"Select Use Profile 성공", 1);
            LogToUI("Cal", $"적용 : {Profile_Select_Index} [{expected_Result}]", 1);

            return true;
        }
        private async Task<bool> Heating_Start()
        {
            int timeout = Device_ReadTimeOut_Value;
            // 1) 패킷 생성
            byte[] tx = new CDCProtocol(Variable.SLAVE, Variable.WRITE, Variable.HEATING_START).GetPacket();

            // 2) 송수신
            byte[] rx = await DevicePort.SendAndReceivePacketAsync(tx, timeout);
            if (rx == null)
            {
                LogToUI("Cal", "응답 없음", 0);
                return false;
            }

            // 3) 응답 검증 및 로그
            string why;
            if (IsEchoAck(tx, rx, out why))
            {
                LogToUI("Cal", "Heating Start 적용 완료", 1);
                return true;
            }
                
            else
            {
                LogToUI("Cal", "Heating Start 적용 실패", 0);
                return false;
            }

        }
        private async Task<bool> Heating_Stop()
        {
            int timeout = Device_ReadTimeOut_Value;
            // 1) 패킷 생성
            byte[] tx = new CDCProtocol(Variable.SLAVE, Variable.WRITE, Variable.HEATING_STOP).GetPacket();

            // 2) 송수신
            byte[] rx = await DevicePort.SendAndReceivePacketAsync(tx, timeout);
            if (rx == null)
            {
                LogToUI("Cal", "응답 없음", 0);
                return false;
            }

            // 3) 응답 검증 및 로그
            string why;
            if (IsEchoAck(tx, rx, out why))
            {
                LogToUI("Cal", "Heating Stop 적용 완료", 1);
                return true;
            }
                
            else
            {
                LogToUI("Cal", "Heating Stop 적용 실패", 0);
                return false;
            }
   
        }
        private async Task<bool> WritePoint3DiffScheduledAsync(CancellationToken ct)
        {
            double? tp, mt;
            lock (_latestLock)
            {
                tp = _lastTempP;
                mt = _lastMaurer;
            }

            if (!tp.HasValue || !mt.HasValue)
            {
                LogToUI("Cal", "[P3] Diff write skip: 값 없음", 0);
                Cal_Result_flag = 0;
                return false;
            }

            float diff = (float)(mt.Value - tp.Value);
            return await WritePoint3DiffAsync(diff, ct);
        }
        private async Task<bool> WritePoint3DiffAsync(float diff, CancellationToken ct)
        {
            byte[] le = BitConverter.GetBytes(diff);
            //if (!BitConverter.IsLittleEndian) Array.Reverse(le);

            byte[] payload = new byte[5 + 4];
            payload[0] = 0x00; payload[1] = 0x12; //쓸 주소
            payload[2] = 0x00; payload[3] = 0x02; //쓸 레지스터 2개
            payload[4] = 0x04; //바이트 수
            payload[5] = le[0];
            payload[6] = le[1];
            payload[7] = le[2];
            payload[8] = le[3]; 

            
            byte[] tx = new CDCProtocol(Variable.SLAVE, Variable.MULTI_WRITE, payload).GetPacket();
            byte[] rx = null;
            try { rx = await DevicePort.SendAndReceivePacketAsync(tx, Device_ReadTimeOut_Value); }
            catch { Cal_Result_flag = 0; }
            

            bool ok = (rx != null) && IsWriteMultiAck(rx, Variable.SLAVE, 0x0012, 0x0002);
            if (!ok) Cal_Result_flag = 0;
            LogToUI("Cal", ok ? string.Format("[P3] Diff {0:F3} write OK", diff)
                      : "[P3] Diff write NG", ok ? 1 : 0);
            return ok;

        }
        private async Task<bool> WritePoint2DiffScheduledAsync(CancellationToken ct)
        {
            double? tp, mt;
            lock (_latestLock)
            {
                tp = _lastTempP;
                mt = _lastMaurer;
            }

            if (!tp.HasValue || !mt.HasValue)
            {
                LogToUI("Cal", "[P2] Diff write skip: 값 없음", 0);
                Cal_Result_flag = 0;
                return false;
            }

            float diff = (float)(mt.Value - tp.Value);
            return await WritePoint2DiffAsync(diff, ct);
        }
        private async Task<bool> WritePoint2DiffAsync(float diff, CancellationToken ct)
        {
            byte[] le = BitConverter.GetBytes(diff);
            //if (!BitConverter.IsLittleEndian) Array.Reverse(le);

            byte[] payload = new byte[5 + 4];
            payload[0] = 0x00; payload[1] = 0x12; //쓸 주소
            payload[2] = 0x00; payload[3] = 0x02; //쓸 레지스터 2개
            payload[4] = 0x04; //바이트 수
            payload[5] = le[0];
            payload[6] = le[1];
            payload[7] = le[2];
            payload[8] = le[3]; 

            
            byte[] tx = new CDCProtocol(Variable.SLAVE, Variable.MULTI_WRITE, payload).GetPacket();
            byte[] rx = null;
            try { rx = await DevicePort.SendAndReceivePacketAsync(tx, Device_ReadTimeOut_Value); }
            catch { Cal_Result_flag = 0; }
            

            bool ok = (rx != null) && IsWriteMultiAck(rx, Variable.SLAVE, 0x0012, 0x0002);
            if (!ok) Cal_Result_flag = 0;
            LogToUI("Cal", ok ? string.Format("[P2] Diff {0:F3} write OK", diff)
                      : "[P2] Diff write NG", ok ? 1 : 0);
            return ok;

        }
        private async Task<bool> WritePoint1DiffScheduledAsync(CancellationToken ct)
        {
            double? tp, mt;
            lock (_latestLock)
            {
                tp = _lastTempP;
                mt = _lastMaurer;
            }

            if (!tp.HasValue || !mt.HasValue)
            {
                LogToUI("Cal", "[P1] Diff write skip: 값 없음", 0);
                Cal_Result_flag = 0;
                return false;
            }

            float diff = (float)(mt.Value - tp.Value);
            return await WritePoint1DiffAsync(diff, ct);
        }
        private async Task<bool> WritePoint1DiffAsync(float diff, CancellationToken ct)
        {
            byte[] le = BitConverter.GetBytes(diff);
            //if (!BitConverter.IsLittleEndian) Array.Reverse(le);

            byte[] payload = new byte[5 + 4];
            payload[0] = 0x00; payload[1] = 0x12; //쓸 주소
            payload[2] = 0x00; payload[3] = 0x02; //쓸 레지스터 2개
            payload[4] = 0x04; //바이트 수
            payload[5] = le[0];
            payload[6] = le[1];
            payload[7] = le[2];
            payload[8] = le[3]; 

            
            byte[] tx = new CDCProtocol(Variable.SLAVE, Variable.MULTI_WRITE, payload).GetPacket();
            byte[] rx = null;
            try { rx = await DevicePort.SendAndReceivePacketAsync(tx, Device_ReadTimeOut_Value); }
            catch { Cal_Result_flag = 0; }
            

            bool ok = (rx != null) && IsWriteMultiAck(rx, Variable.SLAVE, 0x0012, 0x0002);
            if (!ok) Cal_Result_flag = 0;
            LogToUI("Cal", ok ? string.Format("[P1] Diff {0:F3} write OK", diff)
                      : "[P1] Diff write NG", ok ? 1 : 0);

           

            return ok;

        }
        private async Task<bool> ABC_Value_Write(CancellationToken ct)
        {
            int timeout = Device_ReadTimeOut_Value;


            double a = (-X1 * Y2 + X1 * Y3 + X2 * Y1 - X2 * Y3 - X3 * Y1 + X3 * Y2) / ((Math.Pow(X1, 2) * X2) - (Math.Pow(X1, 2) * X3) - (Math.Pow(X2, 2) * X1)
                        + (Math.Pow(X2, 2) * X3) + (Math.Pow(X3, 2) * X1) - (Math.Pow(X3, 2) * X2));

            double b = ((-a * Math.Pow(X2, 2) + a * Math.Pow(X3, 2) + Y2 - Y3)) / (X2 - X3);
            double c = Y1 - a * Math.Pow(X1, 2) - b * X1;

            

            Console.WriteLine($"float a : {(float)a}");
            Console.WriteLine($"float b : {(float)b}");
            Console.WriteLine($"float c : {(float)c}");


            byte[] Byte_a = BitConverter.GetBytes((float)a);
            
            byte[] Byte_b = BitConverter.GetBytes((float)b);
            
            byte[] Byte_c = BitConverter.GetBytes((float)c);
            

            Console.WriteLine($"a : {BitConverter.ToSingle(Byte_a, 0)} Byte => [{BitConverter.ToString(Byte_a)}]");
            Console.WriteLine($"b : {BitConverter.ToSingle(Byte_b, 0)} Byte => [{BitConverter.ToString(Byte_b)}]");
            Console.WriteLine($"c : {BitConverter.ToSingle(Byte_c, 0)} Byte => [{BitConverter.ToString(Byte_c)}]");

            byte[] payload = new byte[5 + 12];
            payload[0] = 0x00; payload[1] = 0x0E;
            payload[2] = 0x00; payload[3] = 0x06;
            payload[4] = 0x0C;
            payload[5] = Byte_a[0]; payload[6] = Byte_a[1]; payload[7] = Byte_a[2]; payload[8] = Byte_a[3]; // a
            payload[9] = Byte_b[0]; payload[10] = Byte_b[1]; payload[11] = Byte_b[2]; payload[12] = Byte_b[3];
            payload[13] = Byte_c[0]; payload[14] = Byte_c[1]; payload[15] = Byte_c[2]; payload[16] = Byte_c[3];



            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.MULTI_WRITE, payload).GetPacket();
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            if (rx1 == null || !IsWriteMultiAck(rx1, Variable.SLAVE, 0x000E, 0x0006))
            {
                LogToUI("Cal", "ABC Write 실패 (멀티쓰기 ACK 불일치)", 0);
                Cal_Result_flag = 0;
                return false;
            }

            return true;
        }
        private async Task<bool> Set_TR_Mode(CancellationToken ct)
        {
            int timeout = Device_ReadTimeOut_Value;
            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.MULTI_WRITE, Variable.SET_TR_MODE_COCONUT).GetPacket();
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            if (rx1 == null || !IsWriteMultiAck(rx1, Variable.SLAVE, 0x0055, 0x000C))
            {
                LogToUI("Cal", "SET TR MODE 실패(멀티 쓰기 ACK 불일치)", 0);
                Cal_Result_flag = 0;
                return false;
            }

            byte[] applyData = new byte[4] { 0x00, 0x6D, 0x01, 0x00 };
            byte[] tx2 = new CDCProtocol(Variable.SLAVE, Variable.WRITE, applyData).GetPacket();
            byte[] rx2 = await DevicePort.SendAndReceivePacketAsync(tx2, timeout);
            if (rx2 == null)
            {
                LogToUI("Cal", "데이터 쓰기 응답 없음", 0);
                Cal_Result_flag = 0;
                return false;
            }

            string why2;
            if (!IsEchoAck(tx2, rx2, out why2))
            {
                LogToUI("Cal", "Set TR Mode 실패", 0);
                Cal_Result_flag = 0;
                return false;
            }



            applyData = new byte[4] { 0x00, 0x6D, 0x00, 0x00 }; //0x6D에 0x0000 write
            byte[] tx3 = new CDCProtocol(Variable.SLAVE, Variable.WRITE, applyData).GetPacket();
            byte[] rx3 = await DevicePort.SendAndReceivePacketAsync(tx3, timeout);
            string why3;
            if (rx3 == null || !IsEchoAck(tx3, rx3, out why3))
            {
                LogToUI("Cal", "Check TR Mode 실패", 0);
                Cal_Result_flag = 0;
                return false;
            }

            byte[] tx4 = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.CHECK_TR_MODE).GetPacket();
            byte[] rx4 = await DevicePort.SendAndReceivePacketAsync(tx4, timeout);
            if (rx4 == null)
            {
                LogToUI("Cal", "Check TR Mode 실패 (Read 응답 없음)", 0);
                Cal_Result_flag = 0;
                return false;
            }

            byte[] frame;
            if (!TryParseAsciiToBytes(rx4, out frame)) frame = rx4;

            if (frame.Length < 3)
            {
                LogToUI("Cal", "Check TR Mode 실패 (응답 프레임 짧음)", 0);
                Cal_Result_flag = 0;
                return false;
            }

            int byteCount = frame[2];
            int payloadOffset = 3;
            if (byteCount != 24 || frame.Length < payloadOffset + byteCount)
            {
                LogToUI("Cal", $"Check TR Mode 실패 (ByteCount=24 예상, 수신={byteCount})", 0);
                Cal_Result_flag = 0;
                return false;
            }

            ushort[] temps = new ushort[12];
            for (int i = 0; i < 12; i++)
            {
                int idx = payloadOffset + i * 2;
                temps[i] = (ushort)(frame[idx] | (frame[idx + 1] << 8));
            }


            ushort[] expTemps = new ushort[12];
            //coconut
            expTemps[3] = 20; expTemps[5] = 7; expTemps[6] = 11; expTemps[7] = 15;

            bool same = true;
            var sb = new StringBuilder();
            sb.AppendLine("TR Mode 값 비교");
            for (int i = 0; i < 12; i++)
            {
                sb.AppendFormat("Point{0}_Temp: 설정={1}, 읽음={2}\r\n", i + 1, expTemps[i], temps[i]);
                if (expTemps[i] != temps[i]) same = false;
            }

            LogToUI("Cal", (same ? "Check TR Mode OK\n" : "Check TR Mode 불일치\n") + sb.ToString(), same ? 1 : 0);
            if (!same) Cal_Result_flag = 0;
            return same;
        }


        private async Task<bool> Set_Calibration_Flag(CancellationToken ct)
        {
            int timeout = Device_ReadTimeOut_Value;
            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.SET_CALIBRATION_FLAG_READ).GetPacket();
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            if (rx1 == null)
            {
                LogToUI("Cal", "Calibration flag read 실패 (프로파일 Read 응답 없음)", 0);
                return false;
            }

            byte[] frame;
            if (!TryParseAsciiToBytes(rx1, out frame))
                frame = rx1;

            if (frame.Length < 3)
            {
                LogToUI("Cal", "Check Profile 실패 (응답 프레임 짧음)", 0);
                return false;
            }

            int byteCount = frame[2];
            int payloadOffset = 3;

            if (frame.Length < payloadOffset + byteCount)
            {
                LogToUI("Cal", "Check Profile 실패 (응답 ByteCount 불일치)", 0);
                return false;
            }

            byte[] Byte_flag_1 = { frame[3], frame[4] };
            byte[] Byte_flag_2 = { frame[5], frame[6] };
            short flag_1 = BitConverter.ToInt16(Byte_flag_1, 0); //Cal?
            short flag_2 = BitConverter.ToInt16(Byte_flag_2, 0);
            Console.WriteLine($"flag 1 : {flag_1}");
            Console.WriteLine($"flag 2 : {flag_2}");
            Console.WriteLine($"Byte_flag_1 : {BitConverter.ToString(Byte_flag_1)}");
            flag_1 = Cal_Result_flag;

            byte[] Byte_flag_1_write = BitConverter.GetBytes(flag_1);
            //Array.Reverse(Byte_flag_1_write);
            //byte[] Byte_flag_2_write = BitConverter.GetBytes(flag_2);
            Console.WriteLine($"Byte_flag_1_Write : {BitConverter.ToString(Byte_flag_1_write)}");
            byte[] payload = new byte[2 + 2];
            payload[0] = 0x00; payload[1] = 0x4B;
            payload[2] = Byte_flag_1_write[0]; payload[3] = Byte_flag_1_write[1];
            
            


            Console.WriteLine($"Now Cal Result Flag : {Cal_Result_flag}");

            byte[] tx2 = new CDCProtocol(Variable.SLAVE, Variable.WRITE, payload).GetPacket();
            byte[] rx2 = await DevicePort.SendAndReceivePacketAsync(tx2, timeout);
            if (rx2 == null)
            {
                LogToUI("Cal", "rx2 응답 없음 [Set Calibration flag]", 0);
                Cal_Result_flag = 0;
                return false;
            }

            string why2;
            if (IsEchoAck(tx2, rx2, out why2))
            {
                LogToUI("Cal", "Set Calibration flag 완료", 1);
                return true;
            }

            else
            {
                LogToUI("Cal", "Set Calibration flag 실패", 0);
                Cal_Result_flag = 0;
                return false;
            }
            
            
        }

     

        private async Task<bool> Set_Verification_Profile(CancellationToken ct)
        {
            if (string.IsNullOrEmpty(Settings.Instance.Profile_File_Verification) || string.IsNullOrEmpty(Settings.Instance.Profile_File_Path_Verification))
            {
                MessageBox.Show("프로파일을 선택해주세요.");
                return false;
            }
            int timeout = Device_ReadTimeOut_Value;
            // 1) Settings.Instance 값으로 데이터 빌드 (온도12 + 시간12 → 총 24워드)
            byte[] data = BuildProfileData(false);
            byte[] payload = new byte[2 + 2 + 1 + data.Length];

            payload[0] = 0x00; payload[1] = 0x55;
            payload[2] = 0x00; payload[3] = 0x18;
            payload[4] = (byte)data.Length;

            Buffer.BlockCopy(data, 0, payload, 5, data.Length);

            // 2) 0x55 ~ 0x6C 연속 쓰기 (Write Multiple Registers)
            //    ↓ CDCProtocol 생성자는 네 프로젝트 규약에 맞게 수정해
            //    예시: (slave, func: WRITE_MULTI, startAddr, payload)
            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.MULTI_WRITE, payload).GetPacket();
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            if (rx1 == null || !IsWriteMultiAck(rx1, Variable.SLAVE, 0x0055, 0x0018))
            {
                LogToUI("Cal", "프로파일 전송 실패(멀티 쓰기 ACK 불일치)", 0);
                Cal_Result_flag = 0;
                return false;
            }

            // 3) 0x6D에 0x0101 쓰기 (프로파일 인덱스/적용 트리거)
            //    예: (slave, func: WRITE, address, data2바이트)
            byte[] applyData = new byte[4] { 0x00, 0x6D, 0x01, 0x01 }; // 0x0101 (big-endian 전송이면 CDCProtocol에서 처리)
            byte[] tx2 = new CDCProtocol(Variable.SLAVE, Variable.WRITE, applyData).GetPacket();
            byte[] rx2 = await DevicePort.SendAndReceivePacketAsync(tx2, timeout);
            if (rx2 == null)
            {
                LogToUI("Cal", "프로파일 적용 응답 없음", 0);
                Cal_Result_flag = 0;
                return false;
            }

            string why2;
            if (IsEchoAck(tx2, rx2, out why2))
            {
                LogToUI("Cal", "Set Verification Profile 완료", 1);
                return true;
            }

            else
            {
                LogToUI("Cal", "Set Verification Profile 실패", 0);
                Cal_Result_flag = 0;
                return false;
            }
        }
        private async Task<bool> Check_Verification_Profile(CancellationToken ct)
        {
            int timeout = Device_ReadTimeOut_Value;
            byte[] applyData = new byte[4] { 0x00, 0x6D, 0x00, 0x01 };
            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.WRITE, applyData).GetPacket();
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            string why1;
            if (rx1 == null || !IsEchoAck(tx1, rx1, out why1))
            {
                LogToUI("Cal", "Check Profile 실패", 0);
                return false;
            }

            byte[] tx2 = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.CHECK_PROFILE).GetPacket();
            byte[] rx2 = await DevicePort.SendAndReceivePacketAsync(tx2, timeout);
            if (rx2 == null)
            {
                LogToUI("Cal", "Check Profile 실패 (프로파일 Read 응답 없음)", 0);
                return false;
            }

            byte[] frame;
            if (!TryParseAsciiToBytes(rx2, out frame))
                frame = rx2;

            if (frame.Length < 3)
            {
                LogToUI("Cal", "Check Profile 실패 (응답 프레임 짧음)", 0);
                return false;
            }

            int byteCount = frame[2];
            int payloadOffset = 3;

            if (frame.Length < payloadOffset + byteCount)
            {
                LogToUI("Cal", "Check Profile 실패 (응답 ByteCount 불일치)", 0);
                return false;
            }

            if (byteCount != 48) // 24워드 * 2바이트
            {
                LogToUI("Cal", $"Check Profile 실패 (예상 ByteCount=48, 수신={byteCount})", 0);
                return false;
            }

            // 기대값 (big-endian → WORD)
            byte[] expected = BuildProfileData(false); // 48 bytes
            ushort[] expectedWords = new ushort[24];
            for (int i = 0; i < 24; i++)
            {
                expectedWords[i] = (ushort)((expected[i * 2] << 8) | expected[i * 2 + 1]);
            }

            // 응답값 (장치가 lo,hi 순서로 보냄 → 리틀엔디언 스왑)
            ushort[] receivedWords = new ushort[24];
            for (int i = 0; i < 24; i++)
            {
                int idx = payloadOffset + i * 2;
                receivedWords[i] = (ushort)(frame[idx] | (frame[idx + 1] << 8)); // lo | (hi<<8)
            }

            bool same = true;
            string tempLog = "온도 값 비교:\r\n";
            string timeLog = "시간 값 비교:\r\n";

            for (int i = 0; i < 12; i++)
            {
                ushort expT = expectedWords[i];        // 기대 온도 i+1
                ushort expTi = expectedWords[12 + i];   // 기대 시간 i+1

                ushort actT = receivedWords[i * 2];     // 읽은 온도 i+1 (짝수 인덱스)
                ushort actTi = receivedWords[i * 2 + 1]; // 읽은 시간 i+1 (홀수 인덱스)

                tempLog += string.Format("Point{0}_Temp: 설정={1}, 읽음={2}\r\n", i + 1, expT, actT);
                timeLog += string.Format("Point{0}_Time: 설정={1}, 읽음={2}\r\n", i + 1, expTi, actTi);

                if (expT != actT || expTi != actTi)
                    same = false;
            }

            LogToUI("Cal",
                (same ? "Check Profile OK\r\n" : "Check Profile 불일치\r\n") + tempLog + timeLog,
                same ? 1 : 0);

            return same;
        }
        private async Task<bool> Flash_Update(CancellationToken ct)
        {
            int timeout = Device_ReadTimeOut_Value;
            // 1) 패킷 생성
            byte[] tx = new CDCProtocol(Variable.SLAVE, Variable.WRITE, Variable.FLASH_UPDATE).GetPacket();

            // 2) 송수신
            byte[] rx = await DevicePort.SendAndReceivePacketAsync(tx, timeout);
            if (rx == null)
            {
                LogToUI("Cal", "응답 없음", 0);
                return false;
            }

            // 3) 응답 검증 및 로그
            string why;
            if (IsEchoAck(tx, rx, out why))
            {
                LogToUI("Cal", "Flash Update 적용 완료", 1);
                return true;
            }

            else
            {
                LogToUI("Cal", "Flash Update 적용 실패", 0);
                return false;
            }
        }

        #endregion

        #region Verification

        private async void PollTimer90_Veri_Tick(object sender, EventArgs e)
        {
            if (_runCts?.IsCancellationRequested == true) return;

            // 1) 게이트 잡고 읽기/표시만 빠르게
            if (!await _ioGate.WaitAsync(0)) return;
            List<ScheduledAction> due = null;

            try
            {
                
                var ct = _runCts?.Token ?? CancellationToken.None;
                bool devOn = DevicePort?.IsOpen == true;
                bool measOn = MeasPort?.IsOpen == true;

                double elapsedSec = _sw.ElapsedMilliseconds / 1000.0;
                var dev = DEV_EMPTY;
                double? maurer = null;

                if (elapsedSec < 330)
                {
                    var devTask = devOn ? ReadDeviceSnapshotAsync(90, ct) : Task.FromResult(DEV_EMPTY);

                    var measTask = measOn ? MeasPort.Ms_ReadTemp(90) : Task.FromResult<double?>(null);

                    await Task.WhenAll(devTask, measTask);

                    dev = devTask.Status == TaskStatus.RanToCompletion ? devTask.Result : DEV_EMPTY;
                    maurer = measTask.Status == TaskStatus.RanToCompletion ? measTask.Result : (double?)null;

                    lock (_latestLock)
                    {
                        if (dev.ok)
                        {
                            _lastTempP = dev.tp; _lastTempA = dev.ta;
                            _lastRunTimeSec = dev.runSec; _lastProfileStep = dev.step;
                        }
                        if (maurer.HasValue) _lastMaurer = maurer.Value;
                    }


                    UpdateUiAndGraph_Veri();
                }

                





                //==================Section 1 Data 산출=====================================================
                if (Settings.Instance.Use_Test_Section_1)
                {
                    if (_S1Active && maurer.HasValue)
                    {
                        //double elapsedSec = _sw.ElapsedMilliseconds / 1000.0;


                        if (elapsedSec >= Settings.Instance.Start_Time_Section_1 && elapsedSec < Settings.Instance.End_Time_Section_1)
                            _S1MaurerTemp.Add(maurer.Value);


                        if (!_S1Averaged && elapsedSec >= Settings.Instance.End_Time_Section_1)
                        {

                            _S1Data = (float)TrimmedMean(_S1MaurerTemp);
                            _S1Averaged = true;
                            _S1Active = false;


                            Console.WriteLine($"S1 Data = {_S1Data}");
                            LogToUI("Veri", $"[Section 1] Samples Count Maurer Temp : {_S1MaurerTemp.Count}", 1);
                            LogToUI("Veri", $"[Section 1] S1 Data = {_S1Data}", 1);
                        }
                    }
                }
                //=================================================================================================

                //==================Section 2 Data 산출=====================================================
                if (Settings.Instance.Use_Test_Section_2)
                {
                    if (_S2Active && maurer.HasValue)
                    {
                        //double elapsedSec = _sw.ElapsedMilliseconds / 1000.0;


                        if (elapsedSec >= Settings.Instance.Start_Time_Section_2 && elapsedSec < Settings.Instance.End_Time_Section_2)
                            _S2MaurerTemp.Add(maurer.Value);


                        if (!_S2Averaged && elapsedSec >= Settings.Instance.End_Time_Section_2)
                        {

                            _S2Data = (float)TrimmedMean(_S2MaurerTemp);
                            _S2Averaged = true;
                            _S2Active = false;


                            Console.WriteLine($"S2 Data = {_S2Data}");
                            LogToUI("Veri", $"[Section 2] Samples Count Maurer Temp : {_S2MaurerTemp.Count}", 1);
                            LogToUI("Veri", $"[Section 2] S2 Data = {_S2Data}", 1);
                        }
                    }
                }
                
                //=================================================================================================

                //==================Section 3 Data 산출=====================================================
                if (Settings.Instance.Use_Test_Section_3)
                {
                    if (_S3Active && maurer.HasValue)
                    {
                        //double elapsedSec = _sw.ElapsedMilliseconds / 1000.0;


                        if (elapsedSec >= Settings.Instance.Start_Time_Section_3 && elapsedSec < Settings.Instance.End_Time_Section_3)
                            _S3MaurerTemp.Add(maurer.Value);


                        if (!_S3Averaged && elapsedSec >= Settings.Instance.End_Time_Section_3)
                        {

                            _S3Data = (float)TrimmedMean(_S3MaurerTemp);
                            _S3Averaged = true;
                            _S3Active = false;


                            Console.WriteLine($"S3 Data = {_S3Data}");
                            LogToUI("Veri", $"[Section 3] Samples Count Maurer Temp : {_S3MaurerTemp.Count}", 1);
                            LogToUI("Veri", $"[Section 3] S3 Data = {_S3Data}", 1);
                        }
                    }
                }
                
                //=================================================================================================

                //==================Section 4 Data 산출=====================================================
                if (Settings.Instance.Use_Test_Section_4)
                {
                    if (_S4Active && maurer.HasValue)
                    {
                        //double elapsedSec = _sw.ElapsedMilliseconds / 1000.0;


                        if (elapsedSec >= Settings.Instance.Start_Time_Section_4 && elapsedSec < Settings.Instance.End_Time_Section_4)
                            _S4MaurerTemp.Add(maurer.Value);


                        if (!_S4Averaged && elapsedSec >= Settings.Instance.End_Time_Section_4)
                        {

                            _S4Data = (float)TrimmedMean(_S4MaurerTemp);
                            _S4Averaged = true;
                            _S4Active = false;


                            Console.WriteLine($"S4 Data = {_S4Data}");
                            LogToUI("Veri", $"[Section 4] Samples Count Maurer Temp : {_S4MaurerTemp.Count}", 1);
                            LogToUI("Veri", $"[Section 4] S4 Data = {_S4Data}", 1);
                        }
                    }
                }
                
                //=================================================================================================

                //==================Section 5 Data 산출=====================================================
                if (Settings.Instance.Use_Test_Section_5)
                {
                    if (_S5Active && maurer.HasValue)
                    {
                        //double elapsedSec = _sw.ElapsedMilliseconds / 1000.0;


                        if (elapsedSec >= Settings.Instance.Start_Time_Section_5 && elapsedSec < Settings.Instance.End_Time_Section_5)
                            _S5MaurerTemp.Add(maurer.Value);


                        if (!_S5Averaged && elapsedSec >= Settings.Instance.End_Time_Section_5)
                        {

                            _S5Data = (float)TrimmedMean(_S5MaurerTemp);
                            _S5Averaged = true;
                            _S5Active = false;


                            Console.WriteLine($"S5 Data = {_S5Data}");
                            LogToUI("Veri", $"[Section 5] Samples Count Maurer Temp : {_S5MaurerTemp.Count}", 1);
                            LogToUI("Veri", $"[Section 5] S5 Data = {_S5Data}", 1);
                        }
                    }
                }
                
                //=================================================================================================

                //==================Section 6 Data 산출=====================================================
                if (Settings.Instance.Use_Test_Section_6)
                {
                    if (_S6Active && maurer.HasValue)
                    {
                        //double elapsedSec = _sw.ElapsedMilliseconds / 1000.0;


                        if (elapsedSec >= Settings.Instance.Start_Time_Section_6 && elapsedSec < Settings.Instance.End_Time_Section_6)
                            _S6MaurerTemp.Add(maurer.Value);


                        if (!_S6Averaged && elapsedSec >= Settings.Instance.End_Time_Section_6)
                        {

                            _S6Data = (float)TrimmedMean(_S6MaurerTemp);
                            _S6Averaged = true;
                            _S6Active = false;


                            Console.WriteLine($"S6 Data = {_S6Data}");
                            LogToUI("Veri", $"[Section 6] Samples Count Maurer Temp : {_S6MaurerTemp.Count}", 1);
                            LogToUI("Veri", $"[Section 6] S6 Data = {_S6Data}", 1);
                        }
                    }
                }
                
                //=================================================================================================

                //==================Section 7 Data 산출=====================================================
                if (Settings.Instance.Use_Test_Section_7)
                {
                    if (_S7Active && maurer.HasValue)
                    {
                        //double elapsedSec = _sw.ElapsedMilliseconds / 1000.0;


                        if (elapsedSec >= Settings.Instance.Start_Time_Section_7 && elapsedSec < Settings.Instance.End_Time_Section_7)
                            _S7MaurerTemp.Add(maurer.Value);


                        if (!_S7Averaged && elapsedSec >= Settings.Instance.End_Time_Section_7)
                        {

                            _S7Data = (float)TrimmedMean(_S7MaurerTemp);
                            _S7Averaged = true;
                            _S7Active = false;


                            Console.WriteLine($"S7 Data = {_S7Data}");
                            LogToUI("Veri", $"[Section 7] Samples Count Maurer Temp : {_S7MaurerTemp.Count}", 1);
                            LogToUI("Veri", $"[Section 7] S7 Data = {_S7Data}", 1);
                        }
                    }
                }
                //=================================================================================================

                //==================Section 8 Data 산출=====================================================
                if (Settings.Instance.Use_Test_Section_8)
                {
                    if (_S8Active && maurer.HasValue)
                    {
                        //double elapsedSec = _sw.ElapsedMilliseconds / 1000.0;


                        if (elapsedSec >= Settings.Instance.Start_Time_Section_8 && elapsedSec < Settings.Instance.End_Time_Section_8)
                            _S8MaurerTemp.Add(maurer.Value);


                        if (!_S8Averaged && elapsedSec >= Settings.Instance.End_Time_Section_8)
                        {

                            _S8Data = (float)TrimmedMean(_S8MaurerTemp);
                            _S8Averaged = true;
                            _S8Active = false;


                            Console.WriteLine($"S8 Data = {_S8Data}");
                            LogToUI("Veri", $"[Section 8] Samples Count Maurer Temp : {_S8MaurerTemp.Count}", 1);
                            LogToUI("Veri", $"[Section 8] S8 Data = {_S8Data}", 1);
                        }
                    }
                }
                
                //=================================================================================================

                //==================Section 9 Data 산출=====================================================
                if (Settings.Instance.Use_Test_Section_9)
                {
                    if (_S9Active && maurer.HasValue)
                    {
                        //double elapsedSec = _sw.ElapsedMilliseconds / 1000.0;


                        if (elapsedSec >= Settings.Instance.Start_Time_Section_9 && elapsedSec < Settings.Instance.End_Time_Section_9)
                            _S9MaurerTemp.Add(maurer.Value);


                        if (!_S9Averaged && elapsedSec >= Settings.Instance.End_Time_Section_9)
                        {

                            _S9Data = (float)TrimmedMean(_S9MaurerTemp);
                            _S9Averaged = true;
                            _S9Active = false;


                            Console.WriteLine($"S9 Data = {_S9Data}");
                            LogToUI("Veri", $"[Section 9] Samples Count Maurer Temp : {_S9MaurerTemp.Count}", 1);
                            LogToUI("Veri", $"[Section 9] S9 Data = {_S9Data}", 1);
                        }
                    }
                }
                
                //=================================================================================================

                //==================Section 10 Data 산출=====================================================
                if (Settings.Instance.Use_Test_Section_10)
                {
                    if (_S10Active && maurer.HasValue)
                    {
                        //double elapsedSec = _sw.ElapsedMilliseconds / 1000.0;


                        if (elapsedSec >= Settings.Instance.Start_Time_Section_10 && elapsedSec < Settings.Instance.End_Time_Section_10)
                            _S10MaurerTemp.Add(maurer.Value);


                        if (!_S10Averaged && elapsedSec >= Settings.Instance.End_Time_Section_10)
                        {

                            _S10Data = (float)TrimmedMean(_S10MaurerTemp);
                            _S10Averaged = true;
                            _S10Active = false;


                            Console.WriteLine($"S10 Data = {_S10Data}");
                            LogToUI("Veri", $"[Section 10] Samples Count Maurer Temp : {_S10MaurerTemp.Count}", 1);
                            LogToUI("Veri", $"[Section 10] S10 Data = {_S10Data}", 1);
                        }
                    }
                }
                
                //=================================================================================================

                //==================Section 11 Data 산출=====================================================
                if (Settings.Instance.Use_Test_Section_11)
                {
                    if (_S11Active && maurer.HasValue)
                    {
                        //double elapsedSec = _sw.ElapsedMilliseconds / 1000.0;


                        if (elapsedSec >= Settings.Instance.Start_Time_Section_11 && elapsedSec < Settings.Instance.End_Time_Section_11)
                            _S11MaurerTemp.Add(maurer.Value);


                        if (!_S11Averaged && elapsedSec >= Settings.Instance.End_Time_Section_11)
                        {

                            _S11Data = (float)TrimmedMean(_S11MaurerTemp);
                            _S11Averaged = true;
                            _S11Active = false;


                            Console.WriteLine($"S11 Data = {_S11Data}");
                            LogToUI("Veri", $"[Section 11] Samples Count Maurer Temp : {_S11MaurerTemp.Count}", 1);
                            LogToUI("Veri", $"[Section 11] S11 Data = {_S11Data}", 1);
                        }
                    }
                }
                
                //=================================================================================================

                //==================Section 12 Data 산출=====================================================
                if (Settings.Instance.Use_Test_Section_12)
                {
                    if (_S12Active && maurer.HasValue)
                    {
                        //double elapsedSec = _sw.ElapsedMilliseconds / 1000.0;


                        if (elapsedSec >= Settings.Instance.Start_Time_Section_12 && elapsedSec < Settings.Instance.End_Time_Section_12)
                            _S12MaurerTemp.Add(maurer.Value);


                        if (!_S12Averaged && elapsedSec >= Settings.Instance.End_Time_Section_12)
                        {

                            _S12Data = (float)TrimmedMean(_S12MaurerTemp);
                            _S12Averaged = true;
                            _S12Active = false;


                            Console.WriteLine($"S12 Data = {_S12Data}");
                            LogToUI("Veri", $"[Section 12] Samples Count Maurer Temp : {_S12MaurerTemp.Count}", 1);
                            LogToUI("Veri", $"[Section 12] S12 Data = {_S12Data}", 1);
                        }
                    }
                }
                
                //=================================================================================================

                // 2) 지금 시각 기준 실행해야 할 스케줄 항목만 복사 (게이트 안에서 큐만 긁어옴)
                long nowMs = _sw.ElapsedMilliseconds;
                due = _schedule.Where(s => !s.Done && nowMs >= s.DueMs).ToList();
                //foreach (var s in due) s.Done = true; // 중복 실행 방지 마킹
            }
            catch (Exception ex)
            {
                LogToUI("Veri", $"Poll error: {ex.Message}", 0);
            }
            finally
            {
                _ioGate.Release(); // <= 폴링/표시 끝났으니 바로 해제
            }

            // 3) 스케줄은 게이트 밖에서, 백그라운드 태스크로 실행
            if (due != null && due.Count > 0)
            {
                foreach (var s in due)
                {
                    _ = Task.Run(async () =>
                    {
                        try
                        {
                            // 액션 내부에서 포트 I/O가 필요하면 s.Action 안에서
                            // await _ioGate.WaitAsync(); try { ... } finally { _ioGate.Release(); }
                            // 패턴으로 잠깐씩만 잡도록 설계
                            await s.Action(_runCts?.Token ?? CancellationToken.None);
                        }
                        catch (Exception ex)
                        {
                            LogToUI("Veri", $"Scheduled action error: {ex.Message}", 0);
                        }
                        finally
                        {
                            s.Done = true;
                        }
                    });
                }
            }

            if (_schedule.Count > 0 && _schedule.All(s => s.Done))
            {
                Console.WriteLine("모든 작업 종료");
                _ = Task.Run(async () => await StopMain_Verification());
                //await StopMain();
            }//다 끝나면 종료
        }

        private void UpdateUiAndGraph_Veri()
        {
            double? tp, ta, mt;
            double run;
            ushort step;

            lock (_latestLock)
            {
                tp = _lastTempP;
                ta = _lastTempA;
                mt = _lastMaurer;
                run = _lastRunTimeSec;
                step = _lastProfileStep;
            }

            void SetIfExists(System.Windows.Forms.Label lbl, string text)
            {
                if (lbl != null) lbl.Text = text;
            }

            // 폼에 실제 있는 컨트롤 이름에 맞춰 호출
            SetIfExists(lblRunTime_Veri, run.ToString("F1"));
            SetIfExists(lblProfileStep_Veri, step.ToString());
            SetIfExists(lblTemp_P_Veri, tp?.ToString("F2") ?? "");
            SetIfExists(lblTemp_A_Veri, ta?.ToString("F2") ?? "");
            SetIfExists(lblMaurerTemp_Veri, mt?.ToString("F2") ?? "");

            // 그래프 X축: Time(x100ms) → ms/100
            double x = _sw.ElapsedMilliseconds / 100.0;

            bool any = false;
            if (tp.HasValue) { _listP_Veri.Add(x, tp.Value); any = true; }
            if (ta.HasValue) { _listA_Veri.Add(x, ta.Value); any = true; }
            if (mt.HasValue) { _listExt_Veri.Add(x, mt.Value); any = true; }

            if (any)
            {
                // 슬라이딩 윈도우 유지 (0~3000)
                var p = TempGraph_Veri.GraphPane;
                if (x > X_WINDOW_VERI)
                {
                    p.XAxis.Scale.Min = x - X_WINDOW_VERI;
                    p.XAxis.Scale.Max = x;
                }
                TempGraph_Veri.AxisChange();
                TempGraph_Veri.Invalidate();
            }
        }

        private async void btnVeriStart_Click(object sender, EventArgs e)
        {
            if (DevicePort == null || !DevicePort.IsOpen)
            {
                MessageBox.Show("디바이스 포트를 연결해주세요");
                return;
            }

            if (MeasPort == null || !MeasPort.IsOpen)
            {
                MessageBox.Show("온도계 포트를 연결해주세요");
                return;
            }

            DevicePort.LogCommToUI = null;
            MeasPort.LogCommToUI = null;

            if (Settings.Instance.Use_Write_Log)
                CsvManager.CsvSave(Environment.NewLine + "=========================== < Verification Start > =============================="
                    + Environment.NewLine);
                  

            bool isPreSettingOk = false;
            isPreSettingOk = await Pre_Veri_Sequence_Task();
            if (!isPreSettingOk) return;

            _taskKind = TaskKind.Verification;

            await Heating_Start_Veri();

            StartMain_Verification();
        }

        private async void btnVeriStop_Click(object sender, EventArgs e)
        {
            deliberate_stop = true;
            await StopMain_Verification();
            deliberate_stop = false;
        }

        private async Task StopMain_Verification()
        {
            _pollTimer90_Veri.Stop();
            await Heating_Stop_Veri();
            if (!deliberate_stop) await SendToDbData(_taskKind);
            _runCts?.Cancel();
            _runCts?.Dispose();
            _runCts = null;           // ← 추가
            _schedule.Clear();        // 선택: 스케줄도 초기화
            DevicePort.LogCommToUI = LogCommToUI;
            MeasPort.LogCommToUI = LogCommToUI;
            if (Settings.Instance.Use_Write_Log)
                CsvManager.CsvSave(Environment.NewLine + "=========================== < Verification End > =============================="
                    + Environment.NewLine);
        }

        private void StartMain_Verification()
        {
            if (_runCts != null) return;              // 이미 동작 중이면 무시
            _runCts = new CancellationTokenSource();

            // 그래프/시간 초기화
            _listA_Veri.Clear(); _listP_Veri.Clear(); _listExt_Veri.Clear();
            _sw.Restart();                             // 기준 시간 0으로
            TempGraph_Veri.AxisChange(); TempGraph_Veri.Invalidate();

            //결과 플래그 초기화
            Veri_Result_flag = 1;

            // 스케줄 구성
            _schedule.Clear();
            //데이터 초기화
            ResetSection1();
            ResetSection2();
            ResetSection3();
            ResetSection4();
            ResetSection5();
            ResetSection6();
            ResetSection7();
            ResetSection8();
            ResetSection9();
            ResetSection10();
            ResetSection11();
            ResetSection12();

            EnqueueAt(330_000, ct => WithIo(() => Jubgement_Pass_Fail(ct), ct)); //Verification 결과 판정
            EnqueueAt(Set_TR_Mode_Time_Veri, ct => WithIo(() => Set_TR_Mode_Veri(ct), ct)); //coconut 모델 적용
            EnqueueAt(Set_Profile_Time_Veri, ct => WithIo(() => Set_Verification_Profile_Veri(ct), ct));
            EnqueueAt(Check_Profile_Time_Veri, ct => WithIo(() => Check_Verification_Profile_Veri(ct), ct));
            EnqueueAt(Set_Flag_Time_Veri, ct => WithIo(() => Set_Verification_Flag_Veri(ct), ct));
            EnqueueAt(Flash_Update_Time_Veri, ct => WithIo(() => Flash_Update_Veri(ct), ct));


            _pollTimer90_Veri.Interval = 90;
            _pollTimer90_Veri.Start();

            //void Enq(long ms, Func<CancellationToken, Task> act) =>
            //    _schedule.Add(new ScheduledAction { DueMs = ms, Action = act, Done = false });
        }
        #endregion

        #region Verification Task Function
        private async Task<bool> Pre_Veri_Sequence_Task()
        {
            bool isOk = false;
            try
            {
                isOk = await Charge_Off_Veri();
                if (!isOk)
                {
                    LogToUI("Veri", "=============================< Charge Off 실패 >=================================", 0);
                    LogToUI("Veri", "=============================< Verification 사전작업 실패 >=================================", 0);
                    return false;
                }


                isOk = await Check_Profile_Verification();
                if (!isOk)
                {
                    LogToUI("Veri", "=============================< Check Profile 실패 >=================================", 0);
                    LogToUI("Veri", "=============================< Verification 사전작업 실패 >=================================", 0);
                    return false;
                }
                isOk = await Check_TR_Mode_Verification();
                if (!isOk)
                {
                    LogToUI("Veri", "=============================< Check TR Mode 실패 >=================================", 0);
                    LogToUI("Veri", "=============================< Verification 사전작업 실패 >=================================", 0);
                    return false;
                }
                
                LogToUI("Veri", "=============================< Verification 사전작업 완료 >=================================", 1);
                return isOk;
            }
            catch (Exception ex)
            {
                LogToUI("Veri", $"{ex.Message}", 0);
                LogToUI("Veri", "=============================< Verification 사전작업 실패 >=================================", 0);
                isOk = false;
                return isOk;
            }
        }
        private async Task<bool> Charge_Off_Veri()
        {
            int timeout = Device_ReadTimeOut_Value;
            // 1) 패킷 생성
            byte[] tx = new CDCProtocol(Variable.SLAVE, Variable.WRITE, Variable.CHARGE_OFF).GetPacket();

            // 2) 송수신
            byte[] rx = await DevicePort.SendAndReceivePacketAsync(tx, timeout);
            if (rx == null)
            {
                LogToUI("Veri", "응답 없음", 0);
                return false;
            }

            // 3) 응답 검증 및 로그
            string why;
            if (IsEchoAck(tx, rx, out why))
            {
                LogToUI("Veri", "Charge OFF 적용 완료", 1);
                return true;
            }

            else
            {
                LogToUI("Veri", "Charge OFF 적용 실패", 0);
                return false;
            }

        }
        private async Task<bool> Heating_Start_Veri()
        {
            int timeout = Device_ReadTimeOut_Value;
            // 1) 패킷 생성
            byte[] tx = new CDCProtocol(Variable.SLAVE, Variable.WRITE, Variable.HEATING_START).GetPacket();

            // 2) 송수신
            byte[] rx = await DevicePort.SendAndReceivePacketAsync(tx, timeout);
            if (rx == null)
            {
                LogToUI("Veri", "응답 없음", 0);
                return false;
            }

            // 3) 응답 검증 및 로그
            string why;
            if (IsEchoAck(tx, rx, out why))
            {
                LogToUI("Veri", "Heating Start 적용 완료", 1);
                return true;
            }

            else
            {
                LogToUI("Veri", "Heating Start 적용 실패", 0);
                return false;
            }

        }
        private async Task<bool> Heating_Stop_Veri()
        {
            int timeout = Device_ReadTimeOut_Value;
            // 1) 패킷 생성
            byte[] tx = new CDCProtocol(Variable.SLAVE, Variable.WRITE, Variable.HEATING_STOP).GetPacket();

            // 2) 송수신
            byte[] rx = await DevicePort.SendAndReceivePacketAsync(tx, timeout);
            if (rx == null)
            {
                LogToUI("Veri", "응답 없음", 0);
                return false;
            }

            // 3) 응답 검증 및 로그
            string why;
            if (IsEchoAck(tx, rx, out why))
            {
                LogToUI("Veri", "Heating Stop 적용 완료", 1);
                return true;
            }

            else
            {
                LogToUI("Veri", "Heating Stop 적용 실패", 0);
                return false;
            }

        }
        private async Task<bool> Check_Profile_Verification()
        {
            int timeout = Device_ReadTimeOut_Value;
            byte[] applyData = new byte[4] { 0x00, 0x6D, 0x00, 0x01 };
            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.WRITE, applyData).GetPacket();
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            string why1;
            if (rx1 == null || !IsEchoAck(tx1, rx1, out why1))
            {
                LogToUI("Veri", "Check Profile 실패", 0);
                return false;
            }

            byte[] tx2 = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.CHECK_PROFILE).GetPacket();
            byte[] rx2 = await DevicePort.SendAndReceivePacketAsync(tx2, timeout);
            if (rx2 == null)
            {
                LogToUI("Veri", "Check Profile 실패 (프로파일 Read 응답 없음)", 0);
                return false;
            }

            byte[] frame;
            if (!TryParseAsciiToBytes(rx2, out frame))
                frame = rx2;

            if (frame.Length < 3)
            {
                LogToUI("Veri", "Check Profile 실패 (응답 프레임 짧음)", 0);
                return false;
            }

            int byteCount = frame[2];
            int payloadOffset = 3;

            if (frame.Length < payloadOffset + byteCount)
            {
                LogToUI("Veri", "Check Profile 실패 (응답 ByteCount 불일치)", 0);
                return false;
            }

            if (byteCount != 48) // 24워드 * 2바이트
            {
                LogToUI("Veri", $"Check Profile 실패 (예상 ByteCount=48, 수신={byteCount})", 0);
                return false;
            }

            // 기대값 (big-endian → WORD)
            byte[] expected = BuildProfileData(false); // 48 bytes
            ushort[] expectedWords = new ushort[24];
            for (int i = 0; i < 24; i++)
            {
                expectedWords[i] = (ushort)((expected[i * 2] << 8) | expected[i * 2 + 1]);
            }

            // 응답값 (장치가 lo,hi 순서로 보냄 → 리틀엔디언 스왑)
            ushort[] receivedWords = new ushort[24];
            for (int i = 0; i < 24; i++)
            {
                int idx = payloadOffset + i * 2;
                receivedWords[i] = (ushort)(frame[idx] | (frame[idx + 1] << 8)); // lo | (hi<<8)
            }

            bool same = true;
            string tempLog = "온도 값 비교:\r\n";
            string timeLog = "시간 값 비교:\r\n";

            for (int i = 0; i < 12; i++)
            {
                ushort expT = expectedWords[i];        // 기대 온도 i+1
                ushort expTi = expectedWords[12 + i];   // 기대 시간 i+1

                ushort actT = receivedWords[i * 2];     // 읽은 온도 i+1 (짝수 인덱스)
                ushort actTi = receivedWords[i * 2 + 1]; // 읽은 시간 i+1 (홀수 인덱스)

                tempLog += string.Format("Point{0}_Temp: 설정={1}, 읽음={2}\r\n", i + 1, expT, actT);
                timeLog += string.Format("Point{0}_Time: 설정={1}, 읽음={2}\r\n", i + 1, expTi, actTi);

                if (expT != actT || expTi != actTi)
                    same = false;
            }

            LogToUI("Veri",
                (same ? "Check Profile OK\r\n" : "Check Profile 불일치\r\n") + tempLog + timeLog,
                same ? 1 : 0);

            return same;
        }

        private async Task<bool> Check_TR_Mode_Verification()
        {
            int timeout = Device_ReadTimeOut_Value;

            byte[] applyData = new byte[4] { 0x00, 0x6D, 0x00, 0x00 }; //0x0100 하니까 안써짐;;
            byte[] tx = new CDCProtocol(Variable.SLAVE, Variable.WRITE, applyData).GetPacket();
            byte[] rx = await DevicePort.SendAndReceivePacketAsync(tx, timeout);
            if (rx == null)
            {
                LogToUI("Veri", "데이터 쓰기 응답 없음", 0);
                Veri_Result_flag = 0;
                return false;
            }

            string why2;
            if (!IsEchoAck(tx, rx, out why2))
            {
                LogToUI("Veri", "Set TR Mode 실패", 0);
                Veri_Result_flag = 0;
                return false;
            }

            byte[] tx4 = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.CHECK_TR_MODE).GetPacket();
            byte[] rx4 = await DevicePort.SendAndReceivePacketAsync(tx4, timeout);
            if (rx4 == null)
            {
                LogToUI("Veri", "Check TR Mode 실패 (Read 응답 없음)", 0);
                Veri_Result_flag = 0;
                return false;
            }

            byte[] frame;
            if (!TryParseAsciiToBytes(rx4, out frame)) frame = rx4;

            if (frame.Length < 3)
            {
                LogToUI("Veri", "Check TR Mode 실패 (응답 프레임 짧음)", 0);
                Veri_Result_flag = 0;
                return false;
            }

            int byteCount = frame[2];
            int payloadOffset = 3;
            if (byteCount != 24 || frame.Length < payloadOffset + byteCount)
            {
                LogToUI("Veri", $"Check TR Mode 실패 (ByteCount=24 예상, 수신={byteCount})", 0);
                Veri_Result_flag = 0;
                return false;
            }

            ushort[] temps = new ushort[12];
            for (int i = 0; i < 12; i++)
            {
                int idx = payloadOffset + i * 2;
                temps[i] = (ushort)(frame[idx] | (frame[idx + 1] << 8));
            }


            ushort[] expTemps = new ushort[12];
            //coconut
            expTemps[3] = 20; expTemps[5] = 7; expTemps[6] = 11; expTemps[7] = 15;

            bool same = true;
            var sb = new StringBuilder();
            sb.AppendLine("TR Mode 값 비교");
            for (int i = 0; i < 12; i++)
            {
                sb.AppendFormat("Point{0}_Temp: 설정={1}, 읽음={2}\r\n", i + 1, expTemps[i], temps[i]);
                if (expTemps[i] != temps[i]) same = false;
            }

            LogToUI("Veri", (same ? "Check TR Mode OK\n" : "Check TR Mode 불일치\n") + sb.ToString(), same ? 1 : 0);
            if (!same) Veri_Result_flag = 0;
            return same;
        }

        

        private async Task<bool> Jubgement_Pass_Fail(CancellationToken ct)
        {
            await Task.Delay(0);
            if (ct.IsCancellationRequested) return false;

            float[] data =
            {
                _S1Data,_S2Data,_S3Data,_S4Data,_S5Data,_S6Data,
                _S7Data,_S8Data,_S9Data,_S10Data,_S11Data,_S12Data
            };

            float[] target =
            {
                (float)Settings.Instance.Point1_Temp_Verification,
                (float)Settings.Instance.Point2_Temp_Verification,
                (float)Settings.Instance.Point3_Temp_Verification,
                (float)Settings.Instance.Point4_Temp_Verification,
                (float)Settings.Instance.Point5_Temp_Verification,
                (float)Settings.Instance.Point6_Temp_Verification,
                (float)Settings.Instance.Point7_Temp_Verification,
                (float)Settings.Instance.Point8_Temp_Verification,
                (float)Settings.Instance.Point9_Temp_Verification,
                (float)Settings.Instance.Point10_Temp_Verification,
                (float)Settings.Instance.Point11_Temp_Verification,
                (float)Settings.Instance.Point12_Temp_Verification
            };

            bool[] use =
            {
                Settings.Instance.Use_Test_Section_1,
                Settings.Instance.Use_Test_Section_2,
                Settings.Instance.Use_Test_Section_3,
                Settings.Instance.Use_Test_Section_4,
                Settings.Instance.Use_Test_Section_5,
                Settings.Instance.Use_Test_Section_6,
                Settings.Instance.Use_Test_Section_7,
                Settings.Instance.Use_Test_Section_8,
                Settings.Instance.Use_Test_Section_9,
                Settings.Instance.Use_Test_Section_10,
                Settings.Instance.Use_Test_Section_11,
                Settings.Instance.Use_Test_Section_12
            };

            float[] th = //임계값
            {
                2.5f, 1.5f, 1.5f, 1.5f, 1.5f, 1.5f,
                1.5f, 1.5f, 1.5f, 1.5f, 1.5f, 1.5f
            };

            StringBuilder sb = new StringBuilder();

            bool allPass = true;

            for (int i = 0; i < 12; i++)
            {
                int secNo = i + 1;

                if (!use[i])
                {
                    // 사용 안 함 → 검사 제외(전체 판정에도 영향 X)
                    sb.AppendLine($" Diff S{secNo} : (skipped) [use=false]");
                    continue;
                }

                // Diff 계산(수집 데이터/타깃 중 NaN/Infinity 방어)
                float diff;
                if (float.IsNaN(data[i]) || float.IsNaN(target[i]) ||
                    float.IsInfinity(data[i]) || float.IsInfinity(target[i]))
                {
                    diff = float.NaN;
                }
                else
                {
                    diff = Math.Abs(data[i] - target[i]);
                }

                // pass 판정: NaN/무한대면 실패로 처리
                bool pass = !(float.IsNaN(diff) || float.IsInfinity(diff)) && (diff < th[i]);

                // 전체 결과 누적
                allPass &= pass;

                // 로그: 숫자 포맷 정리( NaN은 그대로 출력 )
                string diffStr = float.IsNaN(diff) ? "NaN" : diff.ToString();
                sb.AppendLine($" Diff S{secNo} : {diffStr} [< {th[i]} ] {(pass ? "(PASS)" : "(FAIL)")}");
            }

            if (allPass)
            {
                Veri_Result_flag = 1;
                LogToUI("Veri", sb.ToString(), 1);
                LogToUI("Veri", "Verification Result : PASS", 1);
                return true;
            }

            else
            {
                Veri_Result_flag = 0;
                LogToUI("Veri", sb.ToString(), 0);
                LogToUI("Veri", "Verification Result : FAIL", 0);
                return false;
            }

        }

        private async Task<bool> Set_TR_Mode_Veri(CancellationToken ct)
        {
            int timeout = Device_ReadTimeOut_Value;
            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.MULTI_WRITE, Variable.SET_TR_MODE_COCONUT).GetPacket();
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            if (rx1 == null || !IsWriteMultiAck(rx1, Variable.SLAVE, 0x0055, 0x000C))
            {
                LogToUI("Veri", "SET TR MODE 실패(멀티 쓰기 ACK 불일치)", 0);
                Veri_Result_flag = 0;
                return false;
            }

            byte[] applyData = new byte[4] { 0x00, 0x6D, 0x01, 0x00 };
            byte[] tx2 = new CDCProtocol(Variable.SLAVE, Variable.WRITE, applyData).GetPacket();
            byte[] rx2 = await DevicePort.SendAndReceivePacketAsync(tx2, timeout);
            if (rx2 == null)
            {
                LogToUI("Veri", "데이터 쓰기 응답 없음", 0);
                Veri_Result_flag = 0;
                return false;
            }

            string why2;
            if (!IsEchoAck(tx2, rx2, out why2))
            {
                LogToUI("Veri", "Set TR Mode 실패", 0);
                Veri_Result_flag = 0;
                return false;
            }



            applyData = new byte[4] { 0x00, 0x6D, 0x01, 0x00 }; //0x6D에 0x0000 write
            byte[] tx3 = new CDCProtocol(Variable.SLAVE, Variable.WRITE, applyData).GetPacket();
            byte[] rx3 = await DevicePort.SendAndReceivePacketAsync(tx3, timeout);
            string why3;
            if (rx3 == null || !IsEchoAck(tx3, rx3, out why3))
            {
                LogToUI("Veri", "Check TR Mode 실패", 0);
                Veri_Result_flag = 0;
                return false;
            }

            byte[] tx4 = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.CHECK_TR_MODE).GetPacket();
            byte[] rx4 = await DevicePort.SendAndReceivePacketAsync(tx4, timeout);
            if (rx4 == null)
            {
                LogToUI("Veri", "Check TR Mode 실패 (Read 응답 없음)", 0);
                Veri_Result_flag = 0;
                return false;
            }

            byte[] frame;
            if (!TryParseAsciiToBytes(rx4, out frame)) frame = rx4;

            if (frame.Length < 3)
            {
                LogToUI("Veri", "Check TR Mode 실패 (응답 프레임 짧음)", 0);
                Veri_Result_flag = 0;
                return false;
            }

            int byteCount = frame[2];
            int payloadOffset = 3;
            if (byteCount != 24 || frame.Length < payloadOffset + byteCount)
            {
                LogToUI("Veri", $"Check TR Mode 실패 (ByteCount=24 예상, 수신={byteCount})", 0);
                Veri_Result_flag = 0;
                return false;
            }

            ushort[] temps = new ushort[12];
            for (int i = 0; i < 12; i++)
            {
                int idx = payloadOffset + i * 2;
                temps[i] = (ushort)(frame[idx] | (frame[idx + 1] << 8));
            }


            ushort[] expTemps = new ushort[12];
            //coconut
            expTemps[3] = 20; expTemps[5] = 7; expTemps[6] = 11; expTemps[7] = 15;

            bool same = true;
            var sb = new StringBuilder();
            sb.AppendLine("TR Mode 값 비교");
            for (int i = 0; i < 12; i++)
            {
                sb.AppendFormat("Point{0}_Temp: 설정={1}, 읽음={2}\r\n", i + 1, expTemps[i], temps[i]);
                if (expTemps[i] != temps[i]) same = false;
            }

            LogToUI("Veri", (same ? "Check TR Mode OK\n" : "Check TR Mode 불일치\n") + sb.ToString(), same ? 1 : 0);
            if (!same) Veri_Result_flag = 0;
            return same;
        }

        private async Task<bool> Set_Verification_Profile_Veri(CancellationToken ct)
        {
            if (string.IsNullOrEmpty(Settings.Instance.Profile_File_Verification) || string.IsNullOrEmpty(Settings.Instance.Profile_File_Path_Verification))
            {
                MessageBox.Show("프로파일을 선택해주세요.");
                return false;
            }
            int timeout = Device_ReadTimeOut_Value;
            // 1) Settings.Instance 값으로 데이터 빌드 (온도12 + 시간12 → 총 24워드)
            byte[] data = BuildProfileData(false);
            byte[] payload = new byte[2 + 2 + 1 + data.Length];

            payload[0] = 0x00; payload[1] = 0x55;
            payload[2] = 0x00; payload[3] = 0x18;
            payload[4] = (byte)data.Length;

            Buffer.BlockCopy(data, 0, payload, 5, data.Length);

            // 2) 0x55 ~ 0x6C 연속 쓰기 (Write Multiple Registers)
            //    ↓ CDCProtocol 생성자는 네 프로젝트 규약에 맞게 수정해
            //    예시: (slave, func: WRITE_MULTI, startAddr, payload)
            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.MULTI_WRITE, payload).GetPacket();
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            if (rx1 == null || !IsWriteMultiAck(rx1, Variable.SLAVE, 0x0055, 0x0018))
            {
                LogToUI("Veri", "프로파일 전송 실패(멀티 쓰기 ACK 불일치)", 0);
                Veri_Result_flag = 0;
                return false;
            }

            // 3) 0x6D에 0x0101 쓰기 (프로파일 인덱스/적용 트리거)
            //    예: (slave, func: WRITE, address, data2바이트)
            byte[] applyData = new byte[4] { 0x00, 0x6D, 0x01, 0x01 }; // 0x0101 (big-endian 전송이면 CDCProtocol에서 처리)
            byte[] tx2 = new CDCProtocol(Variable.SLAVE, Variable.WRITE, applyData).GetPacket();
            byte[] rx2 = await DevicePort.SendAndReceivePacketAsync(tx2, timeout);
            if (rx2 == null)
            {
                LogToUI("Veri", "프로파일 적용 응답 없음", 0);
                Veri_Result_flag = 0;
                return false;
            }

            string why2;
            if (IsEchoAck(tx2, rx2, out why2))
            {
                LogToUI("Veri", "Set Verification Profile 완료", 1);
                return true;
            }

            else
            {
                LogToUI("Veri", "Set Verification Profile 실패", 0);
                Cal_Result_flag = 0;
                return false;
            }
        }

        private async Task<bool> Check_Verification_Profile_Veri(CancellationToken ct)
        {
            int timeout = Device_ReadTimeOut_Value;
            byte[] applyData = new byte[4] { 0x00, 0x6D, 0x00, 0x01 };
            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.WRITE, applyData).GetPacket();
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            string why1;
            if (rx1 == null || !IsEchoAck(tx1, rx1, out why1))
            {
                LogToUI("Veri", "Check Profile 실패", 0);
                return false;
            }

            byte[] tx2 = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.CHECK_PROFILE).GetPacket();
            byte[] rx2 = await DevicePort.SendAndReceivePacketAsync(tx2, timeout);
            if (rx2 == null)
            {
                LogToUI("Veri", "Check Profile 실패 (프로파일 Read 응답 없음)", 0);
                return false;
            }

            byte[] frame;
            if (!TryParseAsciiToBytes(rx2, out frame))
                frame = rx2;

            if (frame.Length < 3)
            {
                LogToUI("Veri", "Check Profile 실패 (응답 프레임 짧음)", 0);
                return false;
            }

            int byteCount = frame[2];
            int payloadOffset = 3;

            if (frame.Length < payloadOffset + byteCount)
            {
                LogToUI("Veri", "Check Profile 실패 (응답 ByteCount 불일치)", 0);
                return false;
            }

            if (byteCount != 48) // 24워드 * 2바이트
            {
                LogToUI("Veri", $"Check Profile 실패 (예상 ByteCount=48, 수신={byteCount})", 0);
                return false;
            }

            // 기대값 (big-endian → WORD)
            byte[] expected = BuildProfileData(false); // 48 bytes
            ushort[] expectedWords = new ushort[24];
            for (int i = 0; i < 24; i++)
            {
                expectedWords[i] = (ushort)((expected[i * 2] << 8) | expected[i * 2 + 1]);
            }

            // 응답값 (장치가 lo,hi 순서로 보냄 → 리틀엔디언 스왑)
            ushort[] receivedWords = new ushort[24];
            for (int i = 0; i < 24; i++)
            {
                int idx = payloadOffset + i * 2;
                receivedWords[i] = (ushort)(frame[idx] | (frame[idx + 1] << 8)); // lo | (hi<<8)
            }

            bool same = true;
            string tempLog = "온도 값 비교:\r\n";
            string timeLog = "시간 값 비교:\r\n";

            for (int i = 0; i < 12; i++)
            {
                ushort expT = expectedWords[i];        // 기대 온도 i+1
                ushort expTi = expectedWords[12 + i];   // 기대 시간 i+1

                ushort actT = receivedWords[i * 2];     // 읽은 온도 i+1 (짝수 인덱스)
                ushort actTi = receivedWords[i * 2 + 1]; // 읽은 시간 i+1 (홀수 인덱스)

                tempLog += string.Format("Point{0}_Temp: 설정={1}, 읽음={2}\r\n", i + 1, expT, actT);
                timeLog += string.Format("Point{0}_Time: 설정={1}, 읽음={2}\r\n", i + 1, expTi, actTi);

                if (expT != actT || expTi != actTi)
                    same = false;
            }

            LogToUI("Veri",
                (same ? "Check Profile OK\r\n" : "Check Profile 불일치\r\n") + tempLog + timeLog,
                same ? 1 : 0);

            return same;
        }

        private async Task<bool> Set_Verification_Flag_Veri(CancellationToken ct)
        {
            int timeout = Device_ReadTimeOut_Value;
            byte[] tx1 = new CDCProtocol(Variable.SLAVE, Variable.READ, Variable.SET_CALIBRATION_FLAG_READ).GetPacket();
            byte[] rx1 = await DevicePort.SendAndReceivePacketAsync(tx1, timeout);
            if (rx1 == null)
            {
                LogToUI("Veri", "Calibration flag read 실패 (프로파일 Read 응답 없음)", 0);
                return false;
            }

            byte[] frame;
            if (!TryParseAsciiToBytes(rx1, out frame))
                frame = rx1;

            if (frame.Length < 3)
            {
                LogToUI("Veri", "응답 받기 실패 (응답 프레임 짧음)", 0);
                return false;
            }

            int byteCount = frame[2];
            int payloadOffset = 3;

            if (frame.Length < payloadOffset + byteCount)
            {
                LogToUI("Veri", "응답 받기 실패 (응답 ByteCount 불일치)", 0);
                return false;
            }

            byte[] Byte_flag_1 = { frame[3], frame[4] };
            byte[] Byte_flag_2 = { frame[5], frame[6] };
            short flag_1 = BitConverter.ToInt16(Byte_flag_1, 0); //Cal?
            short flag_2 = BitConverter.ToInt16(Byte_flag_2, 0);
            Console.WriteLine($"flag 1 : {flag_1}");
            Console.WriteLine($"flag 2 : {flag_2}");
            Console.WriteLine($"Byte_flag_1 : {BitConverter.ToString(Byte_flag_1)}");
            flag_2 = Veri_Result_flag;

            byte[] Byte_flag_2_write = BitConverter.GetBytes(flag_2);
            //Array.Reverse(Byte_flag_1_write);
            //byte[] Byte_flag_2_write = BitConverter.GetBytes(flag_2);
            Console.WriteLine($"Byte_flag_2_Write : {BitConverter.ToString(Byte_flag_2_write)}");
            byte[] payload = new byte[2 + 2];
            payload[0] = 0x00; payload[1] = 0x4C;
            payload[2] = Byte_flag_2_write[0]; payload[3] = Byte_flag_2_write[1];




            Console.WriteLine($"Now Veri Result Flag : {Veri_Result_flag}");

            byte[] tx2 = new CDCProtocol(Variable.SLAVE, Variable.WRITE, payload).GetPacket();
            byte[] rx2 = await DevicePort.SendAndReceivePacketAsync(tx2, timeout);
            if (rx2 == null)
            {
                LogToUI("Veri", "rx2 응답 없음 [Set Calibration flag]", 0);
                Veri_Result_flag = 0;
                return false;
            }

            string why2;
            if (IsEchoAck(tx2, rx2, out why2))
            {
                LogToUI("Veri", "Set Verification flag 완료", 1);
                return true;
            }

            else
            {
                LogToUI("Veri", "Set Verification flag 실패", 0);
                Veri_Result_flag = 0;
                return false;
            }


        } //Cal 이랑 통틀어서 이거 잘 모르겠음 임시로 만들어둠
        private async Task<bool> Flash_Update_Veri(CancellationToken ct)
        {
            int timeout = Device_ReadTimeOut_Value;
            // 1) 패킷 생성
            byte[] tx = new CDCProtocol(Variable.SLAVE, Variable.WRITE, Variable.FLASH_UPDATE).GetPacket();

            // 2) 송수신
            byte[] rx = await DevicePort.SendAndReceivePacketAsync(tx, timeout);
            if (rx == null)
            {
                LogToUI("Veri", "응답 없음", 0);
                return false;
            }

            // 3) 응답 검증 및 로그
            string why;
            if (IsEchoAck(tx, rx, out why))
            {
                LogToUI("Veri", "Flash Update 적용 완료", 1);
                return true;
            }

            else
            {
                LogToUI("Veri", "Flash Update 적용 실패", 0);
                return false;
            }
        }
        #endregion
    }

   
}
