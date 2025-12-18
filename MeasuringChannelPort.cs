using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;

namespace Heater_Cal_Demo_P4.Communication
{
    public class MeasuringChannelPort : IDisposable
    {

        public int ChannelNo { get; }

        public SerialPort Port { get; private set; }
        public bool IsOpen => Port?.IsOpen ?? false;

        public RichTextBox LogTarget { get; set; }
        public Action<RichTextBox, string, bool> LogCommToUI { get; set; }

        // === 라인 기반 수신 버퍼 ===
        private readonly List<byte> _receiveBuffer = new List<byte>();
        private readonly object _bufferLock = new object();
        private TaskCompletionSource<string> _lineTcs;

        // === 프로토콜 요소 ===
        private const string DevId = "00";       // "00"
        private static readonly Encoding Enc = Encoding.ASCII;
        private const byte CR = 0x0D;

        public MeasuringChannelPort(int ch) => ChannelNo = ch;


        public bool Connect(string portName, int baudRate = 38400)
        {
            try
            {
                Console.WriteLine($"[Measure Serial] CH{ChannelNo} 연결 시도 중: {portName}, BaudRate: {baudRate}");

                if (Port != null)
                {
                    if (Port.IsOpen) Port.Close();
                    Port.Dispose();
                    Port = null;
                }

                Port = new SerialPort(portName, baudRate, Parity.None, 8, StopBits.One)
                {
                    Handshake = Handshake.None,
                    Encoding = Enc,
                    //ReadTimeout = 300,   // 폴링 주기에 맞춰 짧게
                    //WriteTimeout = 300,
                    NewLine = "\r"       // CR 기반 라인 끝
                };

                Port.DataReceived += OnDataReceived;
                Port.Open();
                DiscardBuffers();
                Console.WriteLine($"[Measure Serial] CH{ChannelNo} 연결 완료: {portName}, BaudRate: {baudRate}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Measure Serial] CH{ChannelNo} 연결 실패 예외: {ex.GetType().Name} - {ex.Message}");
                return false;
            }
        }

        public void Disconnect()
        {
            try
            {
                if (IsOpen)
                {
                    Port.DataReceived -= OnDataReceived;
                    Port.Close();
                    Port.Dispose();
                    Port = null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Measure Serial] CH{ChannelNo} Disconnect 예외: {ex.Message}");

            }
        }


        public async Task<bool> SendAsync(byte[] data, int timeoutMs = 200)
        {
            if (!IsOpen) return false;
            using (var cts = new CancellationTokenSource(timeoutMs))
            {
                try
                {
                    await Port.BaseStream.WriteAsync(data, 0, data.Length, cts.Token);
                    if (LogCommToUI != null) LogCommToUI(LogTarget ,BitConverter.ToString(data), true);
                    return true;
                }
                catch (OperationCanceledException)
                {
                    Console.WriteLine($"[Measure Serial] CH{ChannelNo} SendAsync 타임아웃 발생 ({timeoutMs}ms)");
                    return false;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[Measure Serial] CH{ChannelNo} SendAsync 예외: {ex.Message}");
                    return false;
                }
            }

        }

        public void DiscardBuffers()
        {
            try
            {
                Port?.DiscardInBuffer();
                Port?.DiscardOutBuffer();
                lock (_bufferLock) _receiveBuffer.Clear();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void OnDataReceived(object s, SerialDataReceivedEventArgs e)
        {
            try
            {
                lock (_bufferLock)
                {
                    int n = Port.BytesToRead;
                    if (n <= 0) return;

                    byte[] tmp = new byte[n];
                    int read = Port.Read(tmp, 0, n);
                    _receiveBuffer.AddRange(tmp.Take(read));

                    TryExtractLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Measure Serial] CH{ChannelNo} DataRecv 예외: {ex.Message}");
            }
        }

        private void TryExtractLine()
        {
            if (_lineTcs == null || _lineTcs.Task.IsCompleted) return;

            int idx = _receiveBuffer.IndexOf(CR);
            if (idx < 0) return;

            //CR 이전까지가 한 줄
            var lineBytes = _receiveBuffer.Take(idx).ToArray(); // idx + 1 아니냐?
            _receiveBuffer.RemoveRange(0, idx + 1);

            string line = Enc.GetString(lineBytes).Trim();

            if (LogCommToUI != null) LogCommToUI(LogTarget, BitConverter.ToString(Enc.GetBytes(line)), false);

            _lineTcs.TrySetResult(line);
        }

        public async Task<string> SendAndReceiveLineAsync(byte[] request, int timeoutMs)
        {
            DiscardBuffers();
            _lineTcs = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);

            if (!await SendAsync(request)) { _lineTcs.TrySetResult(null); return null; }

            if (LogCommToUI != null ) LogCommToUI(LogTarget, BitConverter.ToString(request), true);

            var timeoutTask = Task.Delay(timeoutMs);
            var completed = await Task.WhenAny(_lineTcs.Task, timeoutTask);
            if (completed == timeoutTask)
            {
                _lineTcs.TrySetCanceled();
                Console.WriteLine($"[Measure Serial] CH{ChannelNo} 응답 타임아웃({timeoutMs}ms)");
                return null;
            }
            string result = await _lineTcs.Task;

            //if (LogCommToUI != null ) LogCommToUI(Enc.GetBytes(result), false);

            return result;
        }

        private static byte[] Build(string cmd2, string param = "")
        {
            if (string.IsNullOrEmpty(cmd2) || cmd2.Length != 2)
                throw new ArgumentException("cmd는 2글자여야 함(ms/em/lk/ce)");

            if (param == null) param = string.Empty;
            string payload = DevId + cmd2 + param;
            byte[] body = Enc.GetBytes(payload);
            byte[] buf = new byte[body.Length + 1];
            Buffer.BlockCopy(body, 0, buf, 0, body.Length);
            buf[buf.Length - 1] = CR;
            return buf;
        }

        public Task<string> Ms_ReadRawAsync(int timeoutMs)
        {
            return SendAndReceiveLineAsync(Build("ms", ""), timeoutMs);
        }

        public async Task<double?> Ms_ReadTemp(int timeoutMs)
        {
            string line = await Ms_ReadRawAsync(timeoutMs);
            int value;
            if (int.TryParse(line, out value))
                return value / 10.0; //10을 왜 나눌까
            return null;
        }

        public Task<string> Em_ReadRawAsync(int timeoutMs)
        {
            return SendAndReceiveLineAsync(Build("em", ""), timeoutMs);
        }

        public async Task<double?> Em_ReadAsync(int timeoutMs)
        {
            string line = await Em_ReadRawAsync(timeoutMs);
            int value;
            if (int.TryParse(line, out value))
                return value / 1000.0;
            return null;
        }

        public Task<string> Lk_GetRawAsync(int timeoutMs)
        {
            return SendAndReceiveLineAsync(Build("lk", ""), timeoutMs);
        }

        public async Task<bool?> Lk_GetAsync(int timeoutMs)
        {
            string line = await Lk_GetRawAsync(timeoutMs);
            if (line == "0") return false;
            if (line == "1") return true;
            return null;
        }

        public Task<string> Lk_SetAsync(bool locked, int timeoutMs)
        {
            return SendAndReceiveLineAsync(Build("lk", locked ? "1" : "0"), timeoutMs);
        }

        public Task<string> Ce_SetAsync(int tenthC, int timeoutMs)
        {
            if (tenthC < 0 || tenthC > 99999) throw new ArgumentOutOfRangeException("tenthC");
            return SendAndReceiveLineAsync(Build("ce", tenthC.ToString("D5")), timeoutMs);
        }

        public void Dispose()
        {
            try { Disconnect(); } catch { }
        }
    }
}
