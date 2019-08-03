using System;
using System.IO;
using System.Diagnostics;
using Gnu.Getopt;

namespace HP3478ACalibration
{
    class Program
    {
        private const string _configFileName = "HP3478ACalibrationSettings.xml";
        private static string _applicationName = null;

        static void Main(string[] args)
        {
            _applicationName = Path.GetFileNameWithoutExtension(Process.GetCurrentProcess().MainModule.FileName);

            // Load configuration file settings
            ApplicationSettings appSettings = null;
            try
            {
                appSettings = new ApplicationSettings(_configFileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: Failed while loading configuration file.\nExtended Error Information:\n{ex.Message ?? "<none>"}");
                Environment.Exit(1);
            }


            // Setup and read application arguments
            LongOpt[] longOpts = new LongOpt[]
            {
                new LongOpt("file", Argument.Required, null, 'f'),      // Calibration file
                new LongOpt("allow-oversize", Argument.No, null, 'o'),  // Allow over-sized calibration files
                new LongOpt("read", Argument.Required, null, 'r'),      // Read calibration data from instrument at given address
                new LongOpt("write", Argument.Required, null, 'w'),     // Write calibration data to instrument at given address
                new LongOpt("help", Argument.No, null, 'h')             // Display help
            };

            Getopt g = new Getopt(_applicationName, args, "f:or:w:h", longOpts)
            {
                Opterr = false  // Do our own error handling
            };

            string optionFile = null;
            bool optionAllowOverSize = false;
            bool optionRead = false;
            bool optionWrite = false;
            int gpibAddress = 0;

            int c;
            while ((c = g.getopt()) != -1)
            {
                switch (c)
                {
                    case 'f':
                        optionFile = g.Optarg;
                        break;

                    case 'o':
                        optionAllowOverSize = true;
                        break;

                    case 'r':
                        optionRead = true;
                        int.TryParse(g.Optarg, out gpibAddress);
                        break;

                    case 'w':
                        optionWrite = true;
                        int.TryParse(g.Optarg, out gpibAddress);
                        break;

                    case 'h':
                        ShowUsage();
                        Environment.Exit(0);
                        break;

                    case '?':
                    default:
                        ShowUsage();
                        Environment.Exit(1);
                        break;
                }
            }

            // Show usage help and exit if:
            //   - Extra parameters are provided.
            //   - No calibration file is specified.
            //   - Both read and write options are set.
            if (g.Optind < args.Length || optionFile == null || (optionRead && optionWrite))
            {
                ShowUsage();
                Environment.Exit(1);
            }

            //Console.WriteLine($"File='{optionFile}'\nOversize={optionAllowOverSize}");
            //Console.WriteLine($"Read={optionRead}\nWrite={optionWrite}\nAddr={gpibAddress}");


            // Validate the calibration file when a write to an instrument will be performed
            // or when only reading a file without connecting to an instrument.
            byte[] fileData = null;
            if (optionWrite || (!optionRead && !optionWrite))
            {
                try
                {
                    FileInfo fi = new FileInfo(optionFile);

                    // Ensure file size is at least 256 bytes
                    if (fi.Length < 256)
                        throw new ApplicationException("File size is less then 256 bytes.");

                    // Ensure file size is exactly 256 bytes if 'allow over-size' option is not specified.
                    if (!optionAllowOverSize && fi.Length != 256)
                        throw new ApplicationException("File size is not 256 bytes.");

                    // Read data from calibration file
                    fileData = ReadFile(optionFile);

                    // Note: The calibration file contains the raw 256 nibbles from the
                    //       instrument's SRAM. When SRAM is read from the meter, 0x40
                    //       is added to the nibble providing an ASCII byte. The file
                    //       contains these bytes.
                    //
                    // SRAM/File Format:
                    //   - The first byte contains the CPU's calibration switch enable check
                    //     value which alternates between 0x40 and 0x4f when the switch is
                    //     enabled. This first byte does not contain calibration data.
                    //
                    //   - The following bytes contain 19 calibration entries each containing
                    //     13 bytes for a total of 247 bytes of calibration data.
                    //
                    //   - The remaining bytes are unused.

                    // Copy raw file data to 2D array for use in calibration library
                    byte[,] calibrationData = new byte[19, 13];
                    Buffer.BlockCopy(fileData, 1, calibrationData, 0, 247);

                    // Validate checksums in calibration file
                    if (!HP3478ACalibration.ValidateData(calibrationData, false))
                        throw new ApplicationException("File contains invalid calibration data.");

                    HP3478ACalibration.PrintData(calibrationData);
                    Console.WriteLine();
                    Console.WriteLine("Calibration file data is valid.");
                    Console.WriteLine();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ERROR: Failed to validate calibration file.\n\nExtended Error Information:\n{ex.Message ?? "<none>"}");
                    Environment.Exit(1);
                }
            }


            if (optionRead)  // Read from instrument
            {
                try
                {
                    // Ensure destination file does not already exist
                    if (File.Exists(optionFile))
                        throw new ApplicationException("Destination file already exists.");

                    Console.WriteLine("Reading calibration data from instrument...");
                    byte[] sramBytes = ReadCalibration(appSettings, gpibAddress);

                    // Copy raw SRAM data to 2D array for use in calibration library
                    byte[,] calibrationData = new byte[19, 13];
                    Buffer.BlockCopy(sramBytes, 1, calibrationData, 0, 247);

                    // Validate checksums in SRAM
                    if (HP3478ACalibration.ValidateData(calibrationData, false))
                        Console.WriteLine("Instrument contains valid calibration data.");
                    else
                        Console.WriteLine("Warning: Instrument contains invalid calibration data.");

                    Console.WriteLine("Writing calibration data to file...");
                    WriteFile(optionFile, sramBytes);
                    Console.WriteLine("Operation complete.");

                    Console.WriteLine();
                    HP3478ACalibration.PrintData(calibrationData);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ERROR: Failed while reading calibration data from instrument.\n\nExtended Error Information:\n{ex.Message ?? "<none>"}");
                    Environment.Exit(1);
                }
            }
            else if (optionWrite)  // Write to instrument
            {
                // Prompt user before writing to instrument
                Console.WriteLine("Warning: This will overwrite all calibration data in the instrument!");

                ConsoleKey keyResponse;
                do
                {
                    Console.Write("Do you want to continue? [y/n] ");
                    keyResponse = Console.ReadKey().Key;
                    Console.WriteLine();

                } while (keyResponse != ConsoleKey.Y && keyResponse != ConsoleKey.N);

                try
                {
                    if (keyResponse == ConsoleKey.Y)
                    {
                        Console.WriteLine("Writing calibration data to instrument...");
                        WriteCalibration(appSettings, gpibAddress, fileData);
                        Console.WriteLine("Operation completed successfully.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ERROR: Failed while writing calibration data to instrument.\n\nExtended Error Information:\n{ex.Message ?? "<none>"}");
                    Environment.Exit(1);
                }
            }
        }

        /// <summary>
        /// Displays application help.
        /// </summary>
        static void ShowUsage()
        {
            Console.WriteLine();
            Console.WriteLine("NAME");
            Console.WriteLine($"    {_applicationName}");

            Console.WriteLine();
            Console.WriteLine("SYNOPSIS");
            Console.WriteLine($"    {_applicationName} -f FILE [-o]");
            Console.WriteLine($"    {_applicationName} -f FILE -r ADDR [-o]");
            Console.WriteLine($"    {_applicationName} -f FILE -w ADDR [-o]");
            Console.WriteLine($"    {_applicationName} -h");

            Console.WriteLine();
            Console.WriteLine("DESCRIPTION");
            Console.WriteLine($"    {_applicationName} reads, writes and verifies HP 3478A meter calibration");
            Console.WriteLine($"    data. Reading and writing of the calibration data is performed using a");
            Console.WriteLine($"    Prologix GPIB-USB compatible adapter.");

            Console.WriteLine();
            Console.WriteLine("OPTIONS");
            Console.WriteLine("    -f FILE, --file=FILE");
            Console.WriteLine("        Use FILE as the calibration file. The calibration file is verified for");
            Console.WriteLine("        correctness when both the -r and -w options are not specified.");

            Console.WriteLine();
            Console.WriteLine("    -o, --allow-oversize");
            Console.WriteLine("        Allow calibration files larger than 256 bytes to be read. (Only the");
            Console.WriteLine("        first 256 bytes of the calibration file will be used when this");
            Console.WriteLine("        option is specified.)");

            Console.WriteLine();
            Console.WriteLine("    -r ADDR, --read=ADDR");
            Console.WriteLine("        Read calibration data from the instrument with GPIB address ADDR");
            Console.WriteLine("        into the calibration file specified by -f. (This option cannot be");
            Console.WriteLine("        specified with option -w.)");

            Console.WriteLine();
            Console.WriteLine("    -w ADDR, --write=ADDR");
            Console.WriteLine("        Write calibration data from the calibration file specified by -f to");
            Console.WriteLine("        the instrument with GPIB address ADDR. (This option cannot be");
            Console.WriteLine("        specified with option -r.)");

            Console.WriteLine();
            Console.WriteLine("    -h, --help");
            Console.WriteLine("        Display this help.");
        }

        /// <summary>
        /// Reads calibration data from an instrument.
        /// </summary>
        /// <param name="appSettings">ApplicationSettings object.</param>
        /// <param name="gpibAddress">GPIB address of instrument.</param>
        /// <returns>256 byte array containing instrument SRAM data.</returns>
        static byte[] ReadCalibration(ApplicationSettings appSettings, int gpibAddress)
        {
            // Connect to GPIB adapter
            GPIBConnection gpibConn = new GPIBConnection(appSettings.SerialPortName,
                appSettings.SerialBaudRate, appSettings.SerialDataBits,
                appSettings.SerialParity, appSettings.SerialStopBits,
                appSettings.SerialFlowControl)
            {
                Timeout = appSettings.GPIBAdapterTimeout,
                GPIBAdapterVersionString = appSettings.GPIBAdapterVersionString,
                GPIBAddress = gpibAddress
            };
            gpibConn.Connect();

            byte[] sramBytes;
            try
            {
                // Test instrument communications
                string cmdResult = gpibConn.QueryInstrument("S");
                if (cmdResult != "0" && cmdResult != "1")
                    throw new ApplicationException("Could not communicate with instrument.");

                // Display "READING CAL" on instrument
                gpibConn.QueryInstrument("D2READING CAL", false);
                gpibConn.QueryInstrumentBinary(new byte[] { 0x1b, 0x0d }, 0);  // Terminate D2 command

                // Setup command buffer to read calibration data.
                //
                // The first byte 'W' is the read SRAM command.
                //
                // The second byte is the ESC (ASCII 27) character used to escape
                // the following byte in order to ensure it gets sent to the
                // instrument and not consumed by the GPIB adapter.
                //
                // The third byte is the SRAM address to be read.
                //
                // Notes:
                //   - SRAM is 256 x 4 (256 4-bit nibbles).
                //
                //   - W<addr> command is used to read SRAM nibbles where <addr>
                //     is an SRAM address from 0-255.
                //
                //   - Nibbles are returned from the instrument with 0x40 added to
                //     them to make them human readable ASCII.
                //     For example, if the nibble value is 0x5, then 0x45 or 'E'
                //     is the byte returned.
                //
                //   - Only escaping CR, LF, ESC and '+' is required, but all
                //     characters are escaped for simplicity.
                byte[] cmdBytes = new byte[3];
                cmdBytes[0] = Convert.ToByte('W');
                cmdBytes[1] = 0x1b;
                cmdBytes[2] = 0x00;

                // Read all 256 nibbles of SRAM
                sramBytes = new byte[256];
                for (int i = 0; i < 256; i++)
                {
                    // Send command to instrument
                    cmdBytes[2] = (byte)i;
                    byte[] recvByte = gpibConn.QueryInstrumentBinary(cmdBytes, 1);

                    if (recvByte == null || recvByte.Length != 1)
                        throw new ApplicationException("Failed while reading calibration data from instrument SRAM.");

                    // Copy received byte to buffer
                    sramBytes[i] = recvByte[0];
                }

                // Return instrument to normal display
                gpibConn.QueryInstrument("D1", false);
            }
            finally
            {
                // Disconnect from GPIB adapter
                gpibConn.Disconnect();
            }

            return sramBytes;
        }

        /// <summary>
        /// Writes calibration data to an instrument.
        /// </summary>
        /// <param name="appSettings">ApplicationSettings object.</param>
        /// <param name="gpibAddress">GPIB address of instrument.</param>
        /// <param name="sramData">256 byte array containing data to write to instrument SRAM.</param>
        static void WriteCalibration(ApplicationSettings appSettings, int gpibAddress, byte[] sramData)
        {
            // Check for valid data buffer
            if (sramData == null || sramData.Length != 256)
                throw new ArgumentException("SRAM data buffer must contain 256 bytes.", "sramData");

            // Connect to GPIB adapter
            GPIBConnection gpibConn = new GPIBConnection(appSettings.SerialPortName,
                appSettings.SerialBaudRate, appSettings.SerialDataBits,
                appSettings.SerialParity, appSettings.SerialStopBits,
                appSettings.SerialFlowControl)
            {
                Timeout = appSettings.GPIBAdapterTimeout,
                GPIBAdapterVersionString = appSettings.GPIBAdapterVersionString,
                GPIBAddress = gpibAddress
            };
            gpibConn.Connect();

            try
            {
                // Test instrument communications
                string cmdResult = gpibConn.QueryInstrument("S");
                if (cmdResult != "0" && cmdResult != "1")
                    throw new ApplicationException("Could not communicate with instrument.");

                // Display "WRITING CAL" on instrument
                gpibConn.QueryInstrument("D2WRITING CAL", false);
                gpibConn.QueryInstrumentBinary(new byte[] { 0x1b, 0x0d }, 0);  // Terminate D2 command

                // Setup command buffer to write calibration data.
                //
                // The first byte 'X' is the write SRAM command.
                //
                // The second and forth bytes are the ESC (ASCII 27) character
                // used to escape the following byte in order to ensure it gets
                // sent to the instrument and not consumed by the GPIB adapter.
                //
                // The third byte is the SRAM address to be written.
                //
                // The fifth byte is the SRAM data to be written.
                //
                // Notes:
                //   - SRAM is 256 x 4 (256 4-bit nibbles).
                //
                //   - X<addr><val> command is used to write SRAM nibbles where
                //     <addr> is an SRAM address from 0-255 and <val> is the data
                //     to be written.
                //
                //   - The instrument ignores the upper four bits of the data byte.
                //     For example, to write the value 0x5 to the instrument,
                //     sending a data byte of 0x05 or 0x45 ('E') are both equivalent.
                //
                //   - Only escaping CR, LF, ESC and '+' is required, but all
                //     characters are escaped for simplicity.
                byte[] writeCmdBytes = new byte[5];
                writeCmdBytes[0] = Convert.ToByte('X');
                writeCmdBytes[1] = 0x1b;
                writeCmdBytes[2] = 0x00;
                writeCmdBytes[3] = 0x1b;
                writeCmdBytes[4] = 0x00;

                // Setup command buffer to read SRAM
                byte[] readCmdBytes = new byte[3];
                readCmdBytes[0] = Convert.ToByte('W');
                readCmdBytes[1] = 0x1b;
                readCmdBytes[2] = 0x00;

                // Write all 256 nibbles of SRAM
                byte[] sramBytes = new byte[256];
                for (int i = 0; i < 256; i++)
                {
                    // Send command to instrument
                    writeCmdBytes[2] = (byte)i;
                    writeCmdBytes[4] = sramData[i];
                    gpibConn.QueryInstrumentBinary(writeCmdBytes, 0);

                    // Read value back for verification of successful write
                    readCmdBytes[2] = (byte)i;
                    byte[] recvByte = gpibConn.QueryInstrumentBinary(readCmdBytes, 1);

                    // Verify correct nibble was read back from instrument
                    if (recvByte == null || recvByte.Length != 1
                        || (recvByte[0] & 0x0f) != (sramData[i] & 0x0f))
                        throw new ApplicationException("Failed while verifying calibration data in instrument SRAM.");
                }

                // Return instrument to normal display
                gpibConn.QueryInstrument("D1", false);
            }
            finally
            {
                // Disconnect from GPIB adapter
                gpibConn.Disconnect();
            }
        }

        /// <summary>
        /// Reads calibration data file.
        /// </summary>
        /// <param name="calibrationFile">Calibration data file.</param>
        /// <returns>256 byte array containing raw instrument SRAM calibration data.</returns>
        static byte[] ReadFile(string calibrationFile)
        {
            // Ensure calibration file contains at least 256 bytes of data
            if (new FileInfo(calibrationFile).Length < 256)
                throw new ApplicationException("File size is less than 256 bytes.");

            // Read 256 bytes from calibration file
            using (BinaryReader br = new BinaryReader(File.OpenRead(calibrationFile)))
            {
                return br.ReadBytes(256);
            }
        }

        /// <summary>
        /// Writes calibration data to file.
        /// </summary>
        /// <param name="calibrationFile">Calibration data file.</param>
        /// <param name="sramData">256 byte array containing data to write to file.</param>
        static void WriteFile(string calibrationFile, byte[] sramData)
        {
            // Check for valid data buffer
            if (sramData == null || sramData.Length != 256)
                throw new ArgumentException("SRAM data buffer must contain 256 bytes.", "sramData");

            File.WriteAllBytes(calibrationFile, sramData);
        }
    }
}
