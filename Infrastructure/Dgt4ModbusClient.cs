using System;
using System.Net.Sockets;
using System.Threading;
using System.Threading.Tasks;
using Modbus.Device;

namespace VUKVWeightApp.Infrastructure;

public sealed class Dgt4ModbusClient : IDisposable
{
        private TcpClient? _tcp;
        private ModbusIpMaster? _master;

        public async Task ConnectAsync(string ip, int port, int timeoutMs)
        {
            Dispose();

            _tcp = new TcpClient
            {
                ReceiveTimeout = timeoutMs,
                SendTimeout = timeoutMs
            };

            await _tcp.ConnectAsync(ip, port);
            _master = ModbusIpMaster.CreateIp(_tcp);
            _master.Transport.Retries = 0;
            _master.Transport.ReadTimeout = timeoutMs;
            _master.Transport.WriteTimeout = timeoutMs;
        }

        public void Dispose()
        {
            try { _master?.Dispose(); } catch { }
            _master = null;

            try { _tcp?.Close(); } catch { }
            _tcp = null;
        }

        private static ushort[] BytesToRegsBE(byte[] bytes)
        {
            if (bytes.Length % 2 != 0) throw new ArgumentException("bytes must be even length");
            var regs = new ushort[bytes.Length / 2];
            for (int i = 0; i < regs.Length; i++)
                regs[i] = (ushort)((bytes[i * 2 + 0] << 8) | bytes[i * 2 + 1]);
            return regs;
        }

        private static int ToInt32_BE(ushort hi, ushort lo) => unchecked((int)(((uint)hi << 16) | lo));
        private static short ToInt16(ushort u) => unchecked((short)u);

        /// <summary>
        /// Change current input page.
        /// DEP.CH uses page 3001 for Gross/Net + status + µV ch1..ch4.
        /// </summary>
        public void ChangePage(byte unitId, uint page)
        {
            // DATA READING (cmd 35 / 0x23)
            WriteCommand(unitId, 0x23, 0);
            Thread.Sleep(120);
            ClearCommandByte(unitId);

            // CHANGE PAGE (cmd 29 / 0x1D)
            WriteCommand(unitId, 0x1D, page);
            Thread.Sleep(160);
            ClearCommandByte(unitId);
        }

        /// <summary>
        /// Zero command (cmd 1 / 0x01). In DEP.CH it's a normal zero of the system.
        /// </summary>
        public void Zero(byte unitId)
        {
            WriteCommand(unitId, 0x01, 0);
            Thread.Sleep(220);
            ClearCommandByte(unitId);
        }

        private void WriteCommand(byte unitId, byte cmd, uint param1)
        {
            if (_master == null) throw new InvalidOperationException("Not connected.");

            // Output Area bytes 0..11 (12 bytes) -> 6 registers
            var outBytes = new byte[12];
            outBytes[0] = 0x00;
            outBytes[1] = cmd;
            outBytes[2] = (byte)((param1 >> 24) & 0xFF);
            outBytes[3] = (byte)((param1 >> 16) & 0xFF);
            outBytes[4] = (byte)((param1 >> 8) & 0xFF);
            outBytes[5] = (byte)(param1 & 0xFF);

            _master.WriteMultipleRegisters(unitId, 0, BytesToRegsBE(outBytes));
        }

        private void ClearCommandByte(byte unitId)
        {
            if (_master == null) return;
            _master.WriteSingleRegister(unitId, 0, 0x0000);
        }

        /// <summary>
        /// DEP.CH page 3001 (0x0BB9):
        /// - Gross weight (int32) at 30001..30002
        /// - Net weight   (int32) at 30003..30004 (not used here)
        /// - InputStatus  (uint16) at 30005
        /// - CommandStatus(uint16) at 30006
        /// - OutputStatus (uint16) at 30007
        /// - SelectedPage (uint16) at 30008
        /// - µV CH1..CH4  (int16)  at 30009..30012
        /// </summary>
        public DepChLive ReadLive3001_DepCh(byte unitId)
        {
            if (_master == null) throw new InvalidOperationException("Not connected.");

            // Read 12 registers: 30001..30012
            var r = _master.ReadInputRegisters(unitId, 0, 12);

            int grossRaw = ToInt32_BE(r[0], r[1]);
            int netRaw = ToInt32_BE(r[2], r[3]);

            ushort inputStatus = r[4];
            ushort commandStatus = r[5];
            ushort outputStatus = r[6];
            ushort selectedPage = r[7];

            short uv1 = ToInt16(r[8]);
            short uv2 = ToInt16(r[9]);
            short uv3 = ToInt16(r[10]);
            short uv4 = ToInt16(r[11]);

            return new DepChLive(grossRaw, netRaw, inputStatus, commandStatus, outputStatus, selectedPage, uv1, uv2, uv3, uv4);
        }

        public readonly record struct DepChLive(
            int GrossRaw,
            int NetRaw,
            ushort InputStatus,
            ushort CommandStatus,
            ushort OutputStatus,
            ushort SelectedPage,
            short Uv1,
            short Uv2,
            short Uv3,
            short Uv4
        );
}
