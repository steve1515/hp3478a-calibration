using System;
using System.Text;
using System.IO;
using System.IO.Ports;
using System.Threading;

namespace HP3478ACalibration
{
    class GPIBConnection
    {
        private readonly SerialPort _serialPort = null;
        private int _timeout = 1000;  // 1 second
        private int _gpibAddress = 1;
        private string _adapterVersionString = "GPIB";


        /// <summary>
        /// Create a GPIB connection instance to a Prologix compatible USB adapter.
        /// </summary>
        /// <param name="portName">Serial port COM port name.</param>
        /// <param name="baudRate">Serial port baud rate setting.</param>
        /// <param name="dataBits">Serial port data bits setting.</param>
        /// <param name="parity">Serial port parity setting.</param>
        /// <param name="stopBits">Serial port stop bits setting.</param>
        /// <param name="flowControl">Serial port flow control setting.</param>
        public GPIBConnection(string portName,
            int baudRate, int dataBits, Parity parity, StopBits stopBits,
            Handshake flowControl)
        {
            _serialPort = new SerialPort()
            {
                PortName = portName,
                BaudRate = baudRate,
                DataBits = dataBits,
                Parity = parity,
                StopBits = stopBits,
                Handshake = flowControl,

                Encoding = Encoding.ASCII,
                NewLine = "\r",

                ReadTimeout = _timeout,
                WriteTimeout = _timeout
            };
        }

        /// <summary>
        /// Internal serial port resource.
        /// </summary>
        public SerialPort SerialPort { get { return _serialPort; } }

        /// <summary>
        /// Timeout used for serial port and GPIB communications.
        /// </summary>
        public int Timeout
        {
            get { return _timeout; }

            set
            {
                if (value < 0)
                    throw new ArgumentOutOfRangeException("Timeout", "Timeout cannot be less than zero.");

                _timeout = value;

                if (_serialPort != null)
                {
                    _serialPort.ReadTimeout = value;
                    _serialPort.WriteTimeout = value;
                }
            }
        }

        /// <summary>
        /// GPIB address of instrument.
        /// </summary>
        public int GPIBAddress
        {
            get { return _gpibAddress; }

            set
            {
                if (value < 1 || value > 30)
                    throw new ArgumentOutOfRangeException("GPIBAddress", "GPIB address must be in the range of 1-30.");

                _gpibAddress = value;

                if (_serialPort != null && _serialPort.IsOpen)
                {
                    _serialPort.WriteLine($"++addr {_gpibAddress}");
                    if (!CheckResponse("++addr", _gpibAddress.ToString()))
                        throw new ApplicationException("Could not set GPIB adapter 'addr' setting.");
                }
            }
        }

        /// <summary>
        /// GPIB adapter version string. (Command '++ver' return value must contain this string.)
        /// </summary>
        public string GPIBAdapterVersionString
        {
            get { return _adapterVersionString; }

            set
            {
                if (string.IsNullOrWhiteSpace(value))
                    throw new ArgumentException("Adapter version string cannot be null or empty.", "GPIBAdapterVersionString");

                _adapterVersionString = value;
            }
        }

        /// <summary>
        /// Connect to GPIB adapter.
        /// </summary>
        public void Connect()
        {
            try
            {
                // Close serial port if already open
                if (_serialPort != null && _serialPort.IsOpen)
                    _serialPort.Close();
            }
            catch (IOException) { }

            _serialPort.Open();
            initGPIBAdapter();
        }

        /// <summary>
        /// Disconnect from GPIB adapter.
        /// </summary>
        public void Disconnect()
        {
            try
            {
                _serialPort?.Close();
            }
            catch (IOException) { }
        }

        /// <summary>
        /// Flush serial port input buffer.
        /// </summary>
        private void flushInputBuffer()
        {
            try
            {
                _serialPort.DiscardInBuffer();
            }
            catch (IOException) { }
            catch (InvalidOperationException) { }
        }

        /// <summary>
        /// Initialize GPIB adapter to good known state.
        /// </summary>
        private void initGPIBAdapter()
        {
            if (_serialPort == null || !_serialPort.IsOpen)
                return;

            // Set and verify GPIB adapter settings
            if (!CheckResponse("++ver", _adapterVersionString, true))
                throw new ApplicationException("Invalid GPIB adapter version string.");

            _serialPort.WriteLine("++mode 1");
            if (!CheckResponse("++mode", "1"))
                throw new ApplicationException("Could not set GPIB adapter 'mode' setting.");

            _serialPort.WriteLine("++auto 0");
            if (!CheckResponse("++auto", "0"))
                throw new ApplicationException("Could not set GPIB adapter 'auto' setting.");

            _serialPort.WriteLine("++eoi 1");
            if (!CheckResponse("++eoi", "1"))
                throw new ApplicationException("Could not set GPIB adapter 'eoi' setting.");

            _serialPort.WriteLine("++eos 0");
            if (!CheckResponse("++eos", "0"))
                throw new ApplicationException("Could not set GPIB adapter 'eos' setting.");

            _serialPort.WriteLine("++eot_enable 0");
            if (!CheckResponse("++eot_enable", "0"))
                throw new ApplicationException("Could not set GPIB adapter 'eot_enable' setting.");

            _serialPort.WriteLine($"++read_tmo_ms {_timeout}");
            if (!CheckResponse("++read_tmo_ms", _timeout.ToString()))
                throw new ApplicationException("Could not set GPIB adapter 'read_tmo_ms' setting.");

            _serialPort.WriteLine("++ifc");

            _serialPort.WriteLine($"++addr {_gpibAddress}");
            if (!CheckResponse("++addr", _gpibAddress.ToString()))
                throw new ApplicationException("Could not set GPIB adapter 'addr' setting.");
        }

        /// <summary>
        /// Sends a command string to the serial port and checks that the expected response is received within the set timeout.
        /// </summary>
        /// <param name="command">Command string to send to serial port.</param>
        /// <param name="expectedResponse">Expected response string.</param>
        /// <param name="anywhereInString">True = Response string may be anywhere in the received data.</param>
        /// <returns>True if expected response was received, False otherwise.</returns>
        public bool CheckResponse(string command, string expectedResponse, bool anywhereInString = false)
        {
            if (_serialPort == null || !_serialPort.IsOpen)
                return false;

            flushInputBuffer();
            _serialPort.WriteLine(command);

            string recvBuffer = string.Empty;
            return SpinWait.SpinUntil(() =>
            {
                if (_serialPort.BytesToRead > 0)
                    recvBuffer += _serialPort.ReadExisting();

                if (anywhereInString && recvBuffer.Contains(expectedResponse))
                    return true;

                if (recvBuffer.StartsWith(expectedResponse, StringComparison.Ordinal))
                    return true;

                return false;
            }, _timeout);
        }

        /// <summary>
        /// Sends a command string to the serial port and returns the resulting string.
        /// </summary>
        /// <param name="command">Command string to send to serial port.</param>
        /// <param name="readResult">True = Read result after sending command; False = Send command only.</param>
        /// <returns>Result string read from serial port.</returns>
        public string QueryInstrument(string command, bool readResult = true)
        {
            if (_serialPort == null || !_serialPort.IsOpen)
                return string.Empty;

            flushInputBuffer();
            _serialPort.WriteLine(command);

            // Don't read back any data if not required
            if (!readResult)
                return string.Empty;

            _serialPort.WriteLine("++read eoi");

            try
            {
                return _serialPort.ReadLine();
            }
            catch (TimeoutException)
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Sends binary data to the serial port and returns the resulting data.
        /// </summary>
        /// <param name="command">Binary data to send to serial port.</param>
        /// <param name="expectedBytes">Number of bytes to read back after sending data.</param>
        /// <returns>Result data read from serial port.</returns>
        public byte[] QueryInstrumentBinary(byte[] command, int expectedBytes)
        {
            if (_serialPort == null || !_serialPort.IsOpen)
                return null;

            flushInputBuffer();
            _serialPort.Write(command, 0, command.Length);
            _serialPort.Write("\r");

            // Don't read back any data if not required
            if (expectedBytes <= 0)
                return null;

            _serialPort.WriteLine("++read eoi");

            try
            {
                // Wait (with timeout) for correct number of bytes in receive buffer
                bool waitResult = SpinWait.SpinUntil(() =>
                {
                    return (_serialPort.BytesToRead >= expectedBytes);
                }, _timeout);

                // Return null if the correct number of bytes hasn't appeared in the receive buffer
                if (!waitResult)
                    return null;

                byte[] buffer = new byte[expectedBytes];
                _serialPort.Read(buffer, 0, expectedBytes);

                return buffer;
            }
            catch (TimeoutException)
            {
                return null;
            }
        }
    }
}
