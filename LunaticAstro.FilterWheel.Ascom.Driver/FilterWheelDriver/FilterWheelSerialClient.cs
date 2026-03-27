using System;
using System.Diagnostics;
using System.IO.Ports;
using System.Threading;
using System.Threading.Tasks;

namespace ASCOM.LunaticAstro.FilterWheel.FilterWheelDriver
{
    internal class FilterWheelSerialClient : IDisposable
    {
        private readonly SerialPort _port;

        public bool IsOpen => _port.IsOpen;

        public Action<string>? Log { get; set; }

        public FilterWheelSerialClient(string portName, int baudRate = 115200)
        {
            _port = new SerialPort(portName, baudRate)
            {
                NewLine = "\n",
                ReadTimeout = 2000,
                WriteTimeout = 2000,
                Handshake = Handshake.None,
                DtrEnable = true,
                RtsEnable = true,
                Encoding = System.Text.Encoding.ASCII
            };
        }

        public void Open()
        {
            if (!_port.IsOpen)
            {
                _port.Open();

                // Arduino resets when port opens
                Thread.Sleep(1500);

                // Clear bootloader noise
                _port.DiscardInBuffer();
                _port.DiscardOutBuffer();

                // Now wait for the wheel to finish its startup routine
                var sw = Stopwatch.StartNew();
                string line;
                int attempts = 0;
                do
                {
                    attempts++;
                    try
                    {
                        line = _port.ReadLine().Trim();

                        // Ignore blank lines or bootloader garbage
                        if (string.IsNullOrWhiteSpace(line) || line.StartsWith("⸮"))
                            continue;

                        // READY signal from firmware
                        if (line.Equals("P1", StringComparison.OrdinalIgnoreCase))
                        {
                            Log?.Invoke("Startup complete: wheel is ready.");
                            break;
                        }

                        // Log anything else for debugging
                        Log?.Invoke($"Startup message: {line}");
                    }
                    catch
                    {
                        // Ignore timeouts during startup
                        Log?.Invoke($"Waiting for startup message {attempts}...");
                    }

                    // Small delay to avoid hammering the port
                    Thread.Sleep(20);

                }
                while (sw.ElapsedMilliseconds < 15000); // allow up to 15 seconds

                if (sw.ElapsedMilliseconds >= 15000)
                    throw new Exception("Filter wheel did not become ready (no P1 received).");
            }
        }

        public void Close()
        {
            if (_port.IsOpen)
                _port.Close();
        }

        // Core synchronous implementation
        public void SendCommand(string command)
        {
            if (!_port.IsOpen)
                throw new InvalidOperationException("Serial port is not open.");

            Log?.Invoke($"TX: {command}");
            _port.WriteLine(command);
        }

        // Async wrapper for existing service code
        public Task SendCommandAsync(string command)
            => Task.Run(() => SendCommand(command));

        public async Task<string> ReadLineAsync(CancellationToken token)
        {
            if (!_port.IsOpen)
                throw new InvalidOperationException("Serial port is not open.");

            return await Task.Run(() =>
            {
                while (true)
                {
                    token.ThrowIfCancellationRequested();

                    try
                    {
                        var line = _port.ReadLine().Trim();
                        if (!string.IsNullOrWhiteSpace(line))
                        {
                            Log?.Invoke($"RX: {line}");
                            return line;
                        }
                    }
                    catch (TimeoutException)
                    {
                        // keep waiting
                    }
                }
            }, token);
        }

        public async Task<string> WaitForResponseAsync(
                Func<string, bool> predicate,
                CancellationToken cancellationToken = default)
        {
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var line = await ReadLineAsync(cancellationToken);

                if (predicate(line))
                    return line;
            }
        }


        public void Dispose()
        {
            if (_port.IsOpen)
                _port.Close();

            _port.Dispose();
        }
    }
}