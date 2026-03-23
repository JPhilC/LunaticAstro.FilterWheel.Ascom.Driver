using ASCOM.DeviceInterface;
using ASCOM.LocalServer;
using ASCOM.LunaticAstro.FilterWheel.FilterWheelDriver.Views;
using ASCOM.Utilities;
using System;
using System.Collections;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace ASCOM.LunaticAstro.FilterWheel.FilterWheelDriver
{
    //[ComVisible(true)]
    //[Guid("4666e34b-23d9-4487-b0fc-1468e1250f44")]
    //[ProgId("ASCOM.LunaticAstroDiyFilterWheel.FilterWheel")]
    //[ServedClassName("ASCOM FilterWheel Driver for LunaticAstroDiyFilterWheel")] // Driver description that appears in the Chooser, customise as required
    //[ClassInterface(ClassInterfaceType.None)]
    public class FilterWheel_New : ReferenceCountedObjectBase, IFilterWheelV3, IDisposable
    {
        private const string driverID = "ASCOM.LunaticAstro.FilterWheel.FilterWheelDriver";
        private const string driverDescription = "Lunatic Astro DIY Filter Wheel";

        private readonly Profile _profile;
        private FilterWheelService _service;
        private bool _connected;
        private int _filterCount;

        // connectedState and connectingState holds the states from this driver instance's perspective, as opposed to the local server's perspective, which may be different because of other client connections.
        internal bool connectedState; // The connected state from this driver's perspective)
        internal bool connectingState; // The connecting state from this driver's perspective)
        internal Exception connectionException = null; // Record any exception thrown if the driver encounters an error when connecting to the hardware using Connect() or Disconnect

        internal TraceLogger tl; // Trace logger object to hold diagnostic information just for this instance of the driver, as opposed to the local server's log, which includes activity from all driver instances.
        private bool disposedValue;

        private Guid uniqueId; // A unique ID for this instance of the driver


        public FilterWheel_New()
        {
            _profile = new Profile { DeviceType = "FilterWheel" };
        }

        #region IFilterWheelV3 implementation

        public bool Connected
        {
            get => _connected;
            set
            {
                if (value == _connected)
                    return;

                if (value)
                {
                    Connect();
                }
                else
                {
                    Disconnect();
                }
            }
        }

        public void Connect()
        {
            if (_connected)
                return;

            string comPort = _profile.GetValue(driverID, "COMPort", string.Empty, "COM1");

            _service = new FilterWheelService(comPort);
            _service.Connect();

            _filterCount = _service.GetFilterSlotCountAsync().Result;

            _connected = true;
        }

        public void Disconnect()
        {
            if (!_connected)
                return;

            _service?.Disconnect();
            _service?.Dispose();
            _service = null;

            _connected = false;
        }

        public int[] FocusOffsets
        {
            get
            {
                if (!_connected)
                    throw new NotConnectedException();

                var offsets = new int[_filterCount];

                for (int i = 0; i < _filterCount; i++)
                {
                    var (_, offset) = _service.GetOffsetForFilterAsync(i + 1).Result;
                    offsets[i] = offset;
                }

                return offsets;
            }
        }

        public string[] Names
        {
            get
            {
                if (!_connected)
                    throw new NotConnectedException();

                var names = new string[_filterCount];

                for (int i = 0; i < _filterCount; i++)
                    names[i] = $"Filter {i + 1}";

                return names;
            }
        }

        public short Position
        {
            get
            {
                if (!_connected)
                    throw new NotConnectedException();

                return (short)_service.GetCurrentPositionAsync().Result;
            }
            set
            {
                if (!_connected)
                    throw new NotConnectedException();

                if (value < 0 || value >= _filterCount)
                    throw new InvalidValueException("Position", value.ToString(), $"0 to {_filterCount - 1}");

                _service.GoToPositionAsync(value).Wait();
            }
        }

        public string Description
        {
            get
            {
                if (!_connected)
                    throw new NotConnectedException();

                var product = _service.GetProductNameAsync().Result;
                var fw = _service.GetFirmwareVersionAsync().Result;
                return $"{driverDescription} - {product} ({fw})";
            }
        }

        public string DriverInfo
            => $"{driverDescription} - ASCOM FilterWheel driver for DIY Nano Filter Wheel";

        public string DriverVersion
            => "1.0";

        public short InterfaceVersion
            => 3;

        public string Name
            => "Lunatic Astro DIY Filter Wheel";

        public ArrayList SupportedActions
            => new ArrayList();

        public bool Connecting
            => false;

        public IStateValueCollection DeviceState
        {
            get
            {
                if (!_connected)
                    throw new NotConnectedException();

                var state = new StateValueCollection
                {
                    { "Position", Position },
                    { "FilterCount", _filterCount }
                };

                return state;
            }
        }

        public void SetupDialog()
        {
            using (var form = new Form())
            {
                form.Text = "DIY Filter Wheel Setup";
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.StartPosition = FormStartPosition.CenterScreen;
                form.Width = 480;
                form.Height = 260;

                var host = new ElementHost
                {
                    Dock = DockStyle.Fill
                };

                var control = new SetupDialogControl();
                host.Child = control;

                form.Controls.Add(host);

                control.CloseRequested += result =>
                {
                    form.DialogResult = result
                        ? DialogResult.OK
                        : DialogResult.Cancel;

                    form.Close();
                };

                form.ShowDialog();
            }
        }

        public string Action(string ActionName, string ActionParameters)
        {
            throw new ActionNotImplementedException(ActionName);
        }

        public void CommandBlind(string Command, bool Raw = false)
        {
            throw new MethodNotImplementedException("CommandBlind is not supported.");
        }

        public bool CommandBool(string Command, bool Raw = false)
        {
            throw new MethodNotImplementedException("CommandBool is not supported.");
        }

        public string CommandString(string Command, bool Raw = false)
        {
            throw new MethodNotImplementedException("CommandString is not supported.");
        }

        public void Dispose()
        {
            if (_connected)
                Disconnect();
        }

        #endregion
    }
}