using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Heater_Cal_Demo_P4.Communication
{
    public class CDCProtocol
    {
        public byte SlaveId { get; }
        public byte Function { get; }
        public byte[] Payload { get; } // Address/Quantity/Value 등 기능별 페이로드
        public byte Lrc { get; }

        public CDCProtocol(byte slaveId, byte function, byte[] payload)
        {
            SlaveId = slaveId;
            Function = function;
            Payload = payload ?? Array.Empty<byte>();

            // LRC는 [SlaveId][Function][Payload...]의 바이너리 합의 2의 보수
            var pdu = new byte[2 + Payload.Length];
            pdu[0] = SlaveId; pdu[1] = Function;
            if (Payload.Length > 0) Buffer.BlockCopy(Payload, 0, pdu, 2, Payload.Length);
            Lrc = ComputeLRC(pdu, 0, pdu.Length);
        }

        // 프레임: ":" + HEX(PDU) + HEX(LRC) + "\r\n"
        public byte[] GetPacket()
        {
            var pdu = new byte[2 + Payload.Length];
            pdu[0] = SlaveId; pdu[1] = Function;
            if (Payload.Length > 0) Buffer.BlockCopy(Payload, 0, pdu, 2, Payload.Length);

            string hex = BytesToHex(pdu) + Lrc.ToString("X2");
            string ascii = ":" + hex + "\r\n";
            return Encoding.ASCII.GetBytes(ascii);
        }

        public string GetPacketAsHexString() => Encoding.ASCII.GetString(GetPacket());

        // ---- Static helpers (편의 빌더) ----
        public static CDCProtocol ReadHolding(byte slaveId, ushort addr, ushort qty)
        {
            byte[] payload = { (byte)(addr >> 8), (byte)(addr & 0xFF), (byte)(qty >> 8), (byte)(qty & 0xFF) };
            return new CDCProtocol(slaveId, 0x03, payload);
        }

        public static CDCProtocol WriteSingle(byte slaveId, ushort addr, ushort value)
        {
            byte[] payload = { (byte)(addr >> 8), (byte)(addr & 0xFF), (byte)(value >> 8), (byte)(value & 0xFF) };
            return new CDCProtocol(slaveId, 0x06, payload);
        }

        public static CDCProtocol WriteMultiple(byte slaveId, ushort addr, ushort[] values)
        {
            if (values == null || values.Length == 0) throw new ArgumentException("values empty");
            ushort qty = (ushort)values.Length;
            byte byteCount = (byte)(qty * 2);

            var payload = new byte[5 + byteCount];
            payload[0] = (byte)(addr >> 8);
            payload[1] = (byte)(addr & 0xFF);
            payload[2] = (byte)(qty >> 8);
            payload[3] = (byte)(qty & 0xFF);
            payload[4] = byteCount;

            int i = 5;
            foreach (var v in values)
            {
                payload[i++] = (byte)(v >> 8);
                payload[i++] = (byte)(v & 0xFF);
            }
            return new CDCProtocol(slaveId, 0x10, payload);
        }

        // ---- Utils ----
        public static byte ComputeLRC(byte[] buf, int offset, int len)
        {
            int sum = 0;
            for (int i = 0; i < len; i++) sum += buf[offset + i];
            return (byte)((-sum) & 0xFF);
        }

        private static string BytesToHex(byte[] bytes)
        {
            var sb = new StringBuilder(bytes.Length * 2);
            foreach (var b in bytes) sb.Append(b.ToString("X2"));
            return sb.ToString();
        }
    }
}

