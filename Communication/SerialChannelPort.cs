using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using Heater_Cal_Demo_P4.Data;
using Heater_Cal_Demo_P4.LIB;

namespace Heater_Cal_Demo_P4.Communication
{
    public class SerialChannelPort
    {
        public int ChannelNo { get; }
        public SerialPort Port { get; private set; }
        public bool IsOpen { get { return Port != null && Port.IsOpen; } }

        public RichTextBox LogTarget { get; set; }
        // true: TX, false: RX
        public Action<RichTextBox, string, bool> LogCommToUI { get; set; }

        private readonly List<byte> _receiveBuffer = new List<byte>();
        private TaskCompletionSource<byte[]> _packetTcs;
        private readonly object _bufferLock = new object();

        // Modbus ASCII framing
        private const byte PacketHeader = 0x3A; // ':'
        private const byte PacketFooter1 = 0x0D; // '\r'
        private const byte PacketFooter2 = 0x0A; // '\n'

        // 요청 간 인터벌(문서: about 90ms) – 여유로 120ms 기본
        private DateTime _lastTxUtc = DateTime.MinValue;
        public int MinIntervalMs { get; set; } = 120;

        public SerialChannelPort(int ch)
        {
            ChannelNo = ch;
        }

        public bool Connect(string portName, int baudRate = 115200)
        {
            try
            {
                Console.WriteLine("[Serial] CH{0} 연결 시도: {1}, {2}-8-N-1", ChannelNo, portName, baudRate);

                if (Port != null)
                {
                    try { if (Port.IsOpen) Port.Close(); } catch { }
                    Port.DataReceived -= OnDataReceived;
                    Port.Dispose();
                    Port = null;
                }

                Port = new SerialPort(portName, baudRate, Parity.None, 8, StopBits.One);
                Port.Encoding = Encoding.ASCII;      // ASCII 라인
                Port.NewLine = "\r\n";               // 줄 끝
                Port.ReadTimeout = 1000;
                Port.WriteTimeout = 1000;

                Port.DataReceived += OnDataReceived;
                Port.Open();

                Port.DiscardInBuffer();
                Port.DiscardOutBuffer();

                Console.WriteLine("[Serial] CH{0} 연결 성공", ChannelNo);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("[Serial] CH{0} 연결 실패: {1} - {2}", ChannelNo, ex.GetType().Name, ex.Message);
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
                Console.WriteLine("[Serial] CH{0} Disconnect 예외: {1}", ChannelNo, ex.Message);
            }
        }

        public async Task<bool> SendAsync(byte[] data, int timeoutMs = 2000)
        {
            if (!IsOpen) return false;

            // 인터벌 보장
            int remain = MinIntervalMs - (int)(DateTime.UtcNow - _lastTxUtc).TotalMilliseconds;
            if (remain > 0) await Task.Delay(remain).ConfigureAwait(false);

            using (var cts = new CancellationTokenSource(timeoutMs))
            {
                try
                {
                    await Port.BaseStream.WriteAsync(data, 0, data.Length, cts.Token).ConfigureAwait(false);
                    _lastTxUtc = DateTime.UtcNow;
                    if (LogCommToUI != null) LogCommToUI(LogTarget, SafeAscii(data), true);
                    if (Settings.Instance.Use_Write_Log)
                    {
                        string Saved_Data = SafeAscii(data);
                        if (!Saved_Data.Contains(":2A030000000DC6")) //90ms마다 데이터 불러오는 건 넣지 않겠다.(너무 많음)
                            CsvManager.CsvSave($"[TX] => {Saved_Data}");

                    }
                    return true;
                }
                catch (OperationCanceledException)
                {
                    Console.WriteLine("[Serial] CH{0} SendAsync 타임아웃({1}ms)", ChannelNo, timeoutMs);
                    return false;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("[Serial] CH{0} SendAsync 예외: {1}", ChannelNo, ex.Message);
                    return false;
                }
            }
        }

        private void OnDataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                lock (_bufferLock)
                {
                    int bytesToRead = Port.BytesToRead;
                    if (bytesToRead <= 0) return;

                    byte[] temp = new byte[bytesToRead];
                    int read = Port.Read(temp, 0, bytesToRead);
                    if (read > 0)
                    {
                        for (int i = 0; i < read; i++) _receiveBuffer.Add(temp[i]);
                    }

                    // 고정 길이 파서 -> ASCII 라인 파서
                    TryExtractFixedPacket();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("[Serial] CH{0} DataReceived 예외: {1}", ChannelNo, ex.Message);
            }
        }

        /// <summary>
        /// 기존 메서드명 유지: 내부는 콜론~CRLF 프레임 추출 + HEX→바이너리 + LRC 검증
        /// </summary>
        private void TryExtractFixedPacket()
        {
            if (_packetTcs == null || _packetTcs.Task.IsCompleted)
                return;

            while (true)
            {
                // 헤더(':') 탐색
                int headerIdx = _receiveBuffer.IndexOf(PacketHeader);
                if (headerIdx < 0)
                {
                    _receiveBuffer.Clear();
                    return;
                }
                if (headerIdx > 0) _receiveBuffer.RemoveRange(0, headerIdx);

                // CRLF 찾기
                int crIdx = -1;
                for (int i = 1; i < _receiveBuffer.Count; i++)
                {
                    if (_receiveBuffer[i] == PacketFooter1)
                    {
                        crIdx = i;
                        break;
                    }
                }
                if (crIdx < 0 || crIdx + 1 >= _receiveBuffer.Count) return; // 아직 라인 미완성

                if (_receiveBuffer[crIdx + 1] != PacketFooter2)
                {
                    // 깨진 라인: ':' 하나 버리고 재시도
                    _receiveBuffer.RemoveAt(0);
                    continue;
                }

                // 프레임(':~CRLF') 확보
                int frameLen = crIdx + 2;
                byte[] frame = new byte[frameLen];
                for (int i = 0; i < frameLen; i++) frame[i] = _receiveBuffer[i];
                _receiveBuffer.RemoveRange(0, frameLen);

                if (frame.Length < 7) continue; // 최소 길이 보호

                // ":" 제외, CRLF 제외 → HEX 본문
                string hex = Encoding.ASCII.GetString(frame, 1, frame.Length - 3);
                if ((hex.Length % 2) != 0) continue;

                byte[] raw = HexToBytes(hex); // [Addr][Func][Data...][LRC]
                if (raw == null || raw.Length < 3) continue;

                // LRC 검증 (마지막 1바이트는 LRC)
                byte rxLrc = raw[raw.Length - 1];
                byte calc = ComputeLRC(raw, 0, raw.Length - 1);
                if (rxLrc != calc)
                {
                    Console.WriteLine("[Serial] CH{0} LRC 불일치: calc={1:X2}, rx={2:X2}", ChannelNo, calc, rxLrc);
                    // 라인 하나는 이미 소비했으니 계속 루프
                    continue;
                }

                if (LogCommToUI != null) LogCommToUI(LogTarget, SafeAscii(frame), false);
                if (Settings.Instance.Use_Write_Log)
                {
                    string Saved_Data = SafeAscii(frame);
                    if (!Saved_Data.Contains(":2A031A")) //90ms마다 데이터 불러오는 건 넣지 않겠다.(너무 많음)
                        CsvManager.CsvSave($"[RX] => {Saved_Data}");
                }
                
                _packetTcs.TrySetResult(frame); // ':'~CRLF 포함 그대로 반환
                return;
            }
        }

        public void DiscardBuffers()
        {
            if (Port == null) return;
            try { Port.DiscardInBuffer(); } catch { }
            try { Port.DiscardOutBuffer(); } catch { }
            lock (_bufferLock) { _receiveBuffer.Clear(); }
        }

        /// <summary>
        /// CDCProtocol.GetPacket()으로 만든 ASCII 프레임을 전송하고,
        /// 응답 ASCII 프레임(':'~CRLF 포함)을 반환.
        /// </summary>
        public async Task<byte[]> SendAndReceivePacketAsync(byte[] data, int timeoutMs)
        {
            DiscardBuffers();

            // VS2019 호환: 옵션 생략(기본 생성자)
            _packetTcs = new TaskCompletionSource<byte[]>();

            bool ok = await SendAsync(data).ConfigureAwait(false);
            if (!ok)
            {
                _packetTcs.TrySetResult(null);
                return null;
            }

            Task timeoutTask = Task.Delay(timeoutMs);
            Task completed = await Task.WhenAny(_packetTcs.Task, timeoutTask).ConfigureAwait(false);
            if (completed == timeoutTask)
            {
                _packetTcs.TrySetCanceled();
                Console.WriteLine("[Serial] CH{0} SendAndReceivePacketAsync 타임아웃 ({1}ms)", ChannelNo, timeoutMs);
                return null;
            }

            if (_packetTcs.Task.IsCompleted && !_packetTcs.Task.IsFaulted)
                return await _packetTcs.Task.ConfigureAwait(false);

            return null;
        }

        // === 기존 보조 메서드 유지 ===
        public bool IsConnected() { return IsOpen; }
        public void DiscardInBuffer() { if (Port != null) Port.DiscardInBuffer(); }
        public void DiscardOutBuffer() { if (Port != null) Port.DiscardOutBuffer(); }

        // ---------- Helpers ----------
        private static string SafeAscii(byte[] buf)
        {
            try { return Encoding.ASCII.GetString(buf); }
            catch { return BitConverter.ToString(buf); }
        }

        private static byte[] HexToBytes(string hex)
        {
            if (string.IsNullOrEmpty(hex) || (hex.Length % 2) != 0) return null;

            int len = hex.Length / 2;
            byte[] dst = new byte[len];
            for (int i = 0; i < len; i++)
            {
                int hi = HexToNib(hex[2 * i]);
                int lo = HexToNib(hex[2 * i + 1]);
                if (hi < 0 || lo < 0) return null;
                dst[i] = (byte)((hi << 4) | lo);
            }
            return dst;
        }

        private static int HexToNib(char c)
        {
            if (c >= '0' && c <= '9') return c - '0';
            char u = (char)(c & ~0x20); // 대문자화
            if (u >= 'A' && u <= 'F') return u - 'A' + 10;
            return -1;
        }

        private static byte ComputeLRC(byte[] buf, int offset, int len)
        {
            int sum = 0;
            for (int i = 0; i < len; i++) sum += buf[offset + i];
            return (byte)((-sum) & 0xFF);
        }
    }
}
