using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASCOM.LunaticAstro.FilterWheel.FilterWheelDriver
{
    internal class FilterWheelSerialClient : IDisposable
    {
        private readonly SerialPort _port;

        public bool IsOpen => _port.IsOpen;

        public FilterWheelSerialClient(string portName, int baudRate = 115200)
        {
            _port = new SerialPort(portName, baudRate)
            {
                NewLine = "\n",
                ReadTimeout = 2000,
                WriteTimeout = 2000
            };
        }

        public void Open()
        {
            if (!_port.IsOpen)
                _port.Open();
        }

        public void Close()
        {
            if (_port.IsOpen)
                _port.Close();
        }

        public async Task<string> SendCommandAsync(string command)
        {
            if (!_port.IsOpen)
                throw new InvalidOperationException("Serial port is not open.");

            // Write command
            _port.WriteLine(command);

            // Read response asynchronously
            return await Task.Run(() => _port.ReadLine().Trim());
        }

        public void Dispose()
        {
            if (_port.IsOpen)
                _port.Close();

            _port.Dispose();
        }
    }

}
