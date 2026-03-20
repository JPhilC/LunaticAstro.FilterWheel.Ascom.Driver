using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ASCOM.LunaticAstro.FilterWheel.FilterWheelDriver
{
    internal class FilterWheelService : IDisposable
    {
        private readonly FilterWheelSerialClient _client;

        public bool IsConnected => _client.IsOpen;

        public FilterWheelService(string comPort, int baudRate = 115200)
        {
            _client = new FilterWheelSerialClient(comPort, baudRate);
        }

        public void Connect()
        {
            if (!IsConnected)
                _client.Open();
        }

        public void Disconnect()
        {
            if (IsConnected)
                _client.Close();
        }

        public void Dispose()
        {
            Disconnect();
            _client.Dispose();
        }

        // ------------------------------------------------------------
        // 1. MOVEMENT COMMANDS
        // ------------------------------------------------------------

        public async Task<int> GoToPositionAsync(int position)
        {
            if (position < 0 || position > 8)
                throw new ArgumentOutOfRangeException(nameof(position));

            var response = await _client.SendCommandAsync($"G{position}");
            return ParsePosition(response);
        }

        public async Task SaveConfigurationAsync()
        {
            var response = await _client.SendCommandAsync("G0");
            if (!response.StartsWith("EEPROM Saved", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException($"Unexpected response to G0: '{response}'");
        }

        // ------------------------------------------------------------
        // 2. OFFSET COMMANDS
        // ------------------------------------------------------------

        public async Task<(int position, int offset)> GetOffsetForFilterAsync(int position)
        {
            if (position < 1 || position > 8)
                throw new ArgumentOutOfRangeException(nameof(position));

            var response = await _client.SendCommandAsync($"O{position}");
            return ParsePositionOffset(response);
        }

        public async Task<(int position, int offset)> IncrementOffsetAsync()
        {
            var response = await _client.SendCommandAsync(")");
            return ParsePositionOffset(response);
        }

        public async Task<(int position, int offset)> DecrementOffsetAsync()
        {
            var response = await _client.SendCommandAsync("(");
            return ParsePositionOffset(response);
        }

        public async Task<(int position, int offset)> SetAbsoluteOffsetAsync(int value)
        {
            var response = await _client.SendCommandAsync($"F{value}");
            return ParsePositionOffset(response);
        }

        // ------------------------------------------------------------
        // 3. INFORMATION COMMANDS (I0–I9)
        // ------------------------------------------------------------

        public async Task<string> GetProductNameAsync()
            => await _client.SendCommandAsync("I0");

        public async Task<string> GetFirmwareVersionAsync()
            => await _client.SendCommandAsync("I1");

        public async Task<int> GetCurrentPositionAsync()
        {
            var response = await _client.SendCommandAsync("I2");
            return ParsePosition(response);
        }

        public async Task<string> GetSerialNumberAsync()
            => await _client.SendCommandAsync("I3");

        public async Task<int> GetMaximumSpeedAsync()
        {
            var response = await _client.SendCommandAsync("I4");
            // "MaxSpeed <maxSpeed>"
            var parts = response.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2 || !int.TryParse(parts[1], out var speed))
                throw new InvalidOperationException($"Unexpected I4 response: '{response}'");
            return speed;
        }

        public async Task<string> GetI5UnusedAsync()
            => await _client.SendCommandAsync("I5");

        public async Task<(int position, int offset)> GetCurrentFilterOffsetAsync()
        {
            var response = await _client.SendCommandAsync("I6");
            return ParsePositionOffset(response);
        }

        public async Task<int> GetThresholdAsync()
        {
            var response = await _client.SendCommandAsync("I7");
            // "Threshold <analogSensorThreshold>"
            var parts = response.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2 || !int.TryParse(parts[1], out var threshold))
                throw new InvalidOperationException($"Unexpected I7 response: '{response}'");
            return threshold;
        }

        public async Task<int> GetFilterSlotCountAsync()
        {
            var response = await _client.SendCommandAsync("I8");
            // "FilterSlots <numberOfFilters>"
            var parts = response.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2 || !int.TryParse(parts[1], out var count))
                throw new InvalidOperationException($"Unexpected I8 response: '{response}'");
            return count;
        }

        public async Task<string> GetI9UnusedAsync()
            => await _client.SendCommandAsync("I9");

        // ------------------------------------------------------------
        // 4. SENSOR TEST COMMANDS (T0–T3)
        // ------------------------------------------------------------

        public async Task<(bool sensor1, bool sensor2)> GetSensorDigitalStateAsync()
        {
            var response = await _client.SendCommandAsync("T0");
            // "Sensors <state> <state>"
            var parts = response.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 3)
                throw new InvalidOperationException($"Unexpected T0/T1 response: '{response}'");

            bool s1 = ParseBool(parts[1]);
            bool s2 = ParseBool(parts[2]);
            return (s1, s2);
        }

        public async Task<bool> IsSensorDigitalAsync()
        {
            var response = await _client.SendCommandAsync("T2");
            // "Digital YES" / "Digital NO"
            var parts = response.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2)
                throw new InvalidOperationException($"Unexpected T2 response: '{response}'");
            return parts[1].Equals("YES", StringComparison.OrdinalIgnoreCase);
        }

        public async Task<bool> IsSensorActiveHighAsync()
        {
            var response = await _client.SendCommandAsync("T3");
            // "Active High YES" / "Active High NO"
            var parts = response.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 3)
                throw new InvalidOperationException($"Unexpected T3 response: '{response}'");
            return parts[2].Equals("YES", StringComparison.OrdinalIgnoreCase);
        }

        // ------------------------------------------------------------
        // 5. RESET / CALIBRATION COMMANDS (R1–R6)
        // ------------------------------------------------------------

        public async Task ReInitialiseAsync()
        {
            await _client.SendCommandAsync("R1");
        }

        public async Task ResetAllOffsetsAsync()
        {
            var response = await _client.SendCommandAsync("R2");
            if (!response.StartsWith("Calibration Removed", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException($"Unexpected R2 response: '{response}'");
        }

        public async Task<string> GetJitterPresetAsync()
            => await _client.SendCommandAsync("R3"); // "Jitter 5"

        public async Task<int> ResetSpeedTo100Async()
        {
            var response = await _client.SendCommandAsync("R4");
            // "MaxSpeed 100%"
            var parts = response.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2)
                throw new InvalidOperationException($"Unexpected R4 response: '{response}'");

            var percentStr = parts[1].TrimEnd('%');
            if (!int.TryParse(percentStr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var percent))
                throw new InvalidOperationException($"Unexpected R4 percent: '{parts[1]}'");

            return percent;
        }

        public async Task<int> ReDetectSensorAndReInitialiseAsync()
        {
            var response = await _client.SendCommandAsync("R5");
            // "Threshold <analogSensorThreshold>"
            var parts = response.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2 || !int.TryParse(parts[1], out var threshold))
                throw new InvalidOperationException($"Unexpected R5 response: '{response}'");
            return threshold;
        }

        public async Task DelayOneSecondAsync()
        {
            // R6 — 1 second delay, no response
            await _client.SendCommandAsync("R6");
        }

        // ------------------------------------------------------------
        // 6. SPEED COMMANDS (Sxxx)
        // ------------------------------------------------------------

        public async Task<int> SetRotationSpeedPercentAsync(int percent)
        {
            if (percent < 0 || percent > 100)
                throw new ArgumentOutOfRangeException(nameof(percent));

            var response = await _client.SendCommandAsync($"S{percent}");
            // "MaxSpeed <percent>%"
            var parts = response.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2)
                throw new InvalidOperationException($"Unexpected S response: '{response}'");

            var percentStr = parts[1].TrimEnd('%');
            if (!int.TryParse(percentStr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result))
                throw new InvalidOperationException($"Unexpected S percent: '{parts[1]}'");

            return result;
        }

        // ------------------------------------------------------------
        // 7. SPOOFED COMMANDS (THRESHOLD/JITTER/PULSE WIDTH)
        // ------------------------------------------------------------

        public async Task<int> AdjustThresholdAsync(bool increase)
        {
            // {0 / }0 — Threshold 0 (spoofed)
            var cmd = increase ? "{0" : "}0";
            var response = await _client.SendCommandAsync(cmd);
            // "Threshold 0"
            var parts = response.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2 || !int.TryParse(parts[1], out var threshold))
                throw new InvalidOperationException($"Unexpected Threshold spoof response: '{response}'");
            return threshold;
        }

        public async Task<int> AdjustJitterAsync(bool increase)
        {
            // [0 / ]0 — Jitter 0 (spoofed)
            var cmd = increase ? "[0" : "]0";
            var response = await _client.SendCommandAsync(cmd);
            // "Jitter 0"
            var parts = response.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2 || !int.TryParse(parts[1], out var jitter))
                throw new InvalidOperationException($"Unexpected Jitter spoof response: '{response}'");
            return jitter;
        }

        public async Task<string> AdjustPulseWidthAsync(bool increase)
        {
            // M0 / N0 — Pulse Width 0uS (spoofed)
            var cmd = increase ? "M0" : "N0";
            var response = await _client.SendCommandAsync(cmd);
            // "Pulse Width 0uS"
            return response;
        }

        // ------------------------------------------------------------
        // PARSING HELPERS
        // ------------------------------------------------------------

        private static int ParsePosition(string response)
        {
            // "P<n>"
            if (!response.StartsWith("P", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException($"Unexpected position response: '{response}'");

            if (!int.TryParse(response.Substring(1), out var pos))
                throw new InvalidOperationException($"Invalid position value in '{response}'");

            return pos;
        }

        private static (int position, int offset) ParsePositionOffset(string response)
        {
            // "P<n> Offset X"
            // Use regex for robustness
            var match = Regex.Match(response, @"^P(\d+)\s+Offset\s+(-?\d+)$",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            if (!match.Success)
                throw new InvalidOperationException($"Unexpected position/offset response: '{response}'");

            var pos = int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
            var offset = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
            return (pos, offset);
        }

        private static bool ParseBool(string value)
        {
            return value.Equals("1", StringComparison.OrdinalIgnoreCase) ||
                   value.Equals("true", StringComparison.OrdinalIgnoreCase) ||
                   value.Equals("YES", StringComparison.OrdinalIgnoreCase);
        }
    }
}