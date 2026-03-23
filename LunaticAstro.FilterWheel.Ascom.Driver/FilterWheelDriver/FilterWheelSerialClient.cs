using System;
using System.IO.Ports;
using System.Threading.Tasks;

internal class FilterWheelSerialClient : IDisposable
{
    private readonly SerialPort _port;

    public bool IsOpen => _port.IsOpen;

    public Action<string>? Log { get; set; }

    public FilterWheelSerialClient(string portName, int baudRate = 9600)
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
            _port.DiscardInBuffer();
            _port.DiscardOutBuffer();
            System.Threading.Thread.Sleep(200); // Allow Arduino reset
        }
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

        Log?.Invoke($"TX: {command}");
        _port.WriteLine(command);

        var response = await Task.Run(() => _port.ReadLine().Trim());
        Log?.Invoke($"RX: {response}");

        return response;
    }

    public string SendCommand(string command)
        => SendCommandAsync(command).GetAwaiter().GetResult();

    public void Dispose()
    {
        if (_port.IsOpen)
            _port.Close();

        _port.Dispose();
    }
}