// TODO fill in this information for your driver, then remove this line!
//
// ASCOM ObservingConditions hardware class for TSOObsCon
//
// Description:	 <To be completed by driver developer>
//
// Implements:	ASCOM ObservingConditions interface version: <To be completed by driver developer>
// Author:		(XXX) Your N. Here <your@email.here>
//

using ASCOM.Astrometry.AstroUtils;
using ASCOM.LocalServer;
using ASCOM.Utilities;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;


namespace ASCOM.TSOObsCon.ObservingConditions
{
    //
    // TODO Replace the not implemented exceptions with code to implement the function or throw the appropriate ASCOM exception.
    //

    /// <summary>
    /// ASCOM ObservingConditions hardware class for TSOObsCon.
    /// </summary>
    [HardwareClass()] // Class attribute flag this as a device hardware class that needs to be disposed by the local server when it exits.
    internal static class ObservingConditionsHardware
    {
        // Constants used for Profile persistence
        internal const string URLProfileName = "URL";
        internal const string URLDefault = "https://tristarobservatory.com/weatherdata/wxdata.txt";
        internal const string traceStateProfileName = "Trace Level";
        internal const string traceStateDefault = "true";

        private static string DriverProgId = ""; // ASCOM DeviceID (COM ProgID) for this driver, the value is set by the driver's class initialiser.
        private static string DriverDescription = ""; // The value is set by the driver's class initialiser.
        internal static string URL; // URL of weather file
        private static bool connectedState; // Local server's connected state
        private static bool runOnce = false; // Flag to enable "one-off" activities only to run once.
        internal static Util utilities; // ASCOM Utilities object for use as required
        internal static AstroUtils astroUtilities; // ASCOM AstroUtilities object for use as required
        internal static TraceLogger tl; // Local server's trace logger object for diagnostic log with information that you specify
        internal static Weather wx;
        private static System.Timers.Timer obsconTimer;

        /// <summary>
        /// Initializes a new instance of the device Hardware class.
        /// </summary>
        static ObservingConditionsHardware()
        {
            try
            {
                // Create the hardware trace logger in the static initialiser.
                // All other initialisation should go in the InitialiseHardware method.
                tl = new TraceLogger("", "TSOObsCon.Hardware");

                // DriverProgId has to be set here because it used by ReadProfile to get the TraceState flag.
                DriverProgId = ObservingConditions.DriverProgId; // Get this device's ProgID so that it can be used to read the Profile configuration values

                // ReadProfile has to go here before anything is written to the log because it loads the TraceLogger enable / disable state.
                ReadProfile(); // Read device configuration from the ASCOM Profile store, including the trace state

                LogMessage("ObservingConditionsHardware", $"Static initialiser completed.");
            }
            catch (Exception ex)
            {
                try { LogMessage("ObservingConditionsHardware", $"Initialisation exception: {ex}"); } catch { }
                MessageBox.Show($"{ex.Message}", "Exception creating ASCOM.TSOObsCon.ObservingConditions", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        /// <summary>
        /// Place device initialisation code here
        /// </summary>
        /// <remarks>Called every time a new instance of the driver is created.</remarks>
        internal static void InitialiseHardware()
        {
            // This method will be called every time a new ASCOM client loads your driver
            LogMessage("InitialiseHardware", $"Start.");

            // Make sure that "one off" activities are only undertaken once
            if (runOnce == false)
            {
                LogMessage("InitialiseHardware", $"Starting one-off initialisation.");

                DriverDescription = ObservingConditions.DriverDescription; // Get this device's Chooser description

                LogMessage("InitialiseHardware", $"ProgID: {DriverProgId}, Description: {DriverDescription}");

                connectedState = false; // Initialise connected to false
                utilities = new Util(); //Initialise ASCOM Utilities object
                astroUtilities = new AstroUtils(); // Initialise ASCOM Astronomy Utilities object

                LogMessage("InitialiseHardware", "Completed basic initialisation");

                // Add your own "one off" device initialisation here e.g. validating existence of hardware and setting up communications
                obsconTimer = new System.Timers.Timer();
                obsconTimer.Interval = 30000;
                obsconTimer.AutoReset = true;
                obsconTimer.Elapsed += new System.Timers.ElapsedEventHandler(OnTimedEvent);
                obsconTimer.Enabled = true;

                wx = new LocalServer.Weather();

                LogMessage("InitialiseHardware", $"One-off initialisation complete.");
                runOnce = true; // Set the flag to ensure that this code is not run again
            }
        }

        // PUBLIC COM INTERFACE IObservingConditions IMPLEMENTATION

        #region Common properties and methods.

        /// <summary>
        /// Displays the Setup Dialogue form.
        /// If the user clicks the OK button to dismiss the form, then
        /// the new settings are saved, otherwise the old values are reloaded.
        /// THIS IS THE ONLY PLACE WHERE SHOWING USER INTERFACE IS ALLOWED!
        /// </summary>
        public static void SetupDialog()
        {
            // Don't permit the setup dialogue if already connected
            if (IsConnected)
                MessageBox.Show("Already connected, just press OK");

            using (SetupDialogForm F = new SetupDialogForm(tl))
            {
                var result = F.ShowDialog();
                if (result == DialogResult.OK)
                {
                    WriteProfile(); // Persist device configuration values to the ASCOM Profile store
                }
            }
        }

        /// <summary>Returns the list of custom action names supported by this driver.</summary>
        /// <value>An ArrayList of strings (SafeArray collection) containing the names of supported actions.</value>
        public static ArrayList SupportedActions
        {
            get
            {
                LogMessage("SupportedActions Get", "Returning empty ArrayList");
                return new ArrayList();
            }
        }

        /// <summary>Invokes the specified device-specific custom action.</summary>
        /// <param name="ActionName">A well known name agreed by interested parties that represents the action to be carried out.</param>
        /// <param name="ActionParameters">List of required parameters or an <see cref="String.Empty">Empty String</see> if none are required.</param>
        /// <returns>A string response. The meaning of returned strings is set by the driver author.
        /// <para>Suppose filter wheels start to appear with automatic wheel changers; new actions could be <c>QueryWheels</c> and <c>SelectWheel</c>. The former returning a formatted list
        /// of wheel names and the second taking a wheel name and making the change, returning appropriate values to indicate success or failure.</para>
        /// </returns>
        public static string Action(string actionName, string actionParameters)
        {
            LogMessage("Action", $"Action {actionName}, parameters {actionParameters} is not implemented");
            throw new ActionNotImplementedException("Action " + actionName + " is not implemented by this driver");
        }

        /// <summary>
        /// Transmits an arbitrary string to the device and does not wait for a response.
        /// Optionally, protocol framing characters may be added to the string before transmission.
        /// </summary>
        /// <param name="Command">The literal command string to be transmitted.</param>
        /// <param name="Raw">
        /// if set to <c>true</c> the string is transmitted 'as-is'.
        /// If set to <c>false</c> then protocol framing characters may be added prior to transmission.
        /// </param>
        public static void CommandBlind(string command, bool raw)
        {
            CheckConnected("CommandBlind");
            // TODO The optional CommandBlind method should either be implemented OR throw a MethodNotImplementedException
            // If implemented, CommandBlind must send the supplied command to the mount and return immediately without waiting for a response

            throw new MethodNotImplementedException($"CommandBlind - Command:{command}, Raw: {raw}.");
        }

        /// <summary>
        /// Transmits an arbitrary string to the device and waits for a boolean response.
        /// Optionally, protocol framing characters may be added to the string before transmission.
        /// </summary>
        /// <param name="Command">The literal command string to be transmitted.</param>
        /// <param name="Raw">
        /// if set to <c>true</c> the string is transmitted 'as-is'.
        /// If set to <c>false</c> then protocol framing characters may be added prior to transmission.
        /// </param>
        /// <returns>
        /// Returns the interpreted boolean response received from the device.
        /// </returns>
        public static bool CommandBool(string command, bool raw)
        {
            CheckConnected("CommandBool");
            // TODO The optional CommandBool method should either be implemented OR throw a MethodNotImplementedException
            // If implemented, CommandBool must send the supplied command to the mount, wait for a response and parse this to return a True or False value

            throw new MethodNotImplementedException($"CommandBool - Command:{command}, Raw: {raw}.");
        }

        /// <summary>
        /// Transmits an arbitrary string to the device and waits for a string response.
        /// Optionally, protocol framing characters may be added to the string before transmission.
        /// </summary>
        /// <param name="Command">The literal command string to be transmitted.</param>
        /// <param name="Raw">
        /// if set to <c>true</c> the string is transmitted 'as-is'.
        /// If set to <c>false</c> then protocol framing characters may be added prior to transmission.
        /// </param>
        /// <returns>
        /// Returns the string response received from the device.
        /// </returns>
        public static string CommandString(string command, bool raw)
        {
            CheckConnected("CommandString");
            // TODO The optional CommandString method should either be implemented OR throw a MethodNotImplementedException
            // If implemented, CommandString must send the supplied command to the mount and wait for a response before returning this to the client

            throw new MethodNotImplementedException($"CommandString - Command:{command}, Raw: {raw}.");
        }

        /// <summary>
        /// Deterministically release both managed and unmanaged resources that are used by this class.
        /// </summary>
        /// <remarks>
        /// TODO: Release any managed or unmanaged resources that are used in this class.
        /// 
        /// Do not call this method from the Dispose method in your driver class.
        ///
        /// This is because this hardware class is decorated with the <see cref="HardwareClassAttribute"/> attribute and this Dispose() method will be called 
        /// automatically by the  local server executable when it is irretrievably shutting down. This gives you the opportunity to release managed and unmanaged 
        /// resources in a timely fashion and avoid any time delay between local server close down and garbage collection by the .NET runtime.
        ///
        /// For the same reason, do not call the SharedResources.Dispose() method from this method. Any resources used in the static shared resources class
        /// itself should be released in the SharedResources.Dispose() method as usual. The SharedResources.Dispose() method will be called automatically 
        /// by the local server just before it shuts down.
        /// 
        /// </remarks>
        public static void Dispose()
        {
            try { LogMessage("Dispose", $"Disposing of assets and closing down."); } catch { }

            try
            {
                // Clean up the trace logger and utility objects
                tl.Enabled = false;
                tl.Dispose();
                tl = null;
            }
            catch { }

            try
            {
                utilities.Dispose();
                utilities = null;
            }
            catch { }

            try
            {
                astroUtilities.Dispose();
                astroUtilities = null;
            }
            catch { }
        }

        /// <summary>
        /// Set True to connect to the device hardware. Set False to disconnect from the device hardware.
        /// You can also read the property to check whether it is connected. This reports the current hardware state.
        /// </summary>
        /// <value><c>true</c> if connected to the hardware; otherwise, <c>false</c>.</value>
        public static bool Connected
        {
            get
            {
                LogMessage("Connected", $"Get {IsConnected}");
                return IsConnected;
            }
            set
            {
                LogMessage("Connected", $"Set {value}");
                if (value == IsConnected)
                    return;

                if (value)
                {
                    LogMessage("Connected Set", $"Connected");
                    connectedState = true;
                    wx.checkWeather(URL, tl);
                }
                else
                {
                    LogMessage("Connected Set", $"Disconnected");
                    connectedState = false;
                }
            }
        }

        /// <summary>
        /// Returns a description of the device, such as manufacturer and model number. Any ASCII characters may be used.
        /// </summary>
        /// <value>The description.</value>
        public static string Description
        {
            // TODO customise this device description if required
            get
            {
                LogMessage("Description Get", DriverDescription);
                return DriverDescription;
            }
        }

        /// <summary>
        /// Descriptive and version information about this ASCOM driver.
        /// </summary>
        public static string DriverInfo
        {
            get
            {
                string driverInfo = "Data from " + URL;
                LogMessage("DriverInfo Get", driverInfo);
                return driverInfo;
            }
        }

        /// <summary>
        /// A string containing only the major and minor version of the driver formatted as 'm.n'.
        /// </summary>
        public static string DriverVersion
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                string driverVersion = $"{version.Major}.{version.Minor}.{version.Revision}.{version.Build}";
                LogMessage("DriverVersion Get", driverVersion);
                return driverVersion;
            }
        }

        /// <summary>
        /// The interface version number that this device supports.
        /// </summary>
        public static short InterfaceVersion
        {
            // set by the driver wizard
            get
            {
                LogMessage("InterfaceVersion Get", "1");
                return Convert.ToInt16("1");
            }
        }

        /// <summary>
        /// The short name of the driver, for display purposes
        /// </summary>
        public static string Name
        {
            // TODO customise this device name as required
            get
            {
                string name = "Short driver name - please customise";
                LogMessage("Name Get", name);
                return name;
            }
        }

        #endregion

        #region IObservingConditions Implementation

        // Time and wind speed values
        private static Dictionary<DateTime, double> winds = new Dictionary<DateTime, double>();

        /// <summary>
        /// Gets and sets the time period over which observations wil be averaged
        /// </summary>
        internal static double AveragePeriod
        {
            get
            {
                LogMessage("AveragePeriod", "get - 0");
                return 0.0;
            }
            set
            {
                LogMessage("AveragePeriod Set", value.ToString());
                if (value != 0.0)
                {
                    LogMessage("AveragePeriod Set", "Invalid value, set to 0");
                    throw new InvalidValueException("AveragePeriod", value.ToString(), "0 only");
                }
            }
        }

        /// <summary>
        /// Amount of sky obscured by cloud
        /// </summary>
        internal static double CloudCover
        {
            get
            {
                double cloudcon;
                switch (wx.CloudCondition)
                {
                    case 1:
                        {
                            cloudcon = 0.0d;
                            break;
                        }
                    case 2:
                        {
                            cloudcon = 50.0d;
                            break;
                        }
                    case 3:
                        {
                            cloudcon = 100.0d;
                            break;
                        }
                    default:
                        {
                            cloudcon = 100.0d;
                            break;
                        }
                }
                LogMessage("CloudCover", cloudcon.ToString());
                return cloudcon;
            }
        }

        /// <summary>
        /// Atmospheric dew point at the observatory in deg C
        /// </summary>
        internal static double DewPoint
        {
            get
            {
                // Calculaton of dewpoint for wx file doesn't seem to be working.  "Temporary" cheat here that I'm sure will never be replaced.
                double dp;
                double T = wx.Temp;
                double RH = wx.Hum;
                dp = 243.04d * (Math.Log(RH / 100d) + 17.625d * T / (243.04d + T)) / (17.625d - Math.Log(RH / 100d) - 17.625d * T / (243.04d + T));
                LogMessage("DewPoint", dp.ToString());
                return Math.Round(dp, 2);
            }
        }

        /// <summary>
        /// Atmospheric relative humidity at the observatory in percent
        /// </summary>
        internal static double Humidity
        {
            get
            {
                LogMessage("Humidity", wx.Hum.ToString());
                return wx.Hum;
            }
        }

        /// <summary>
        /// Atmospheric pressure at the observatory in hectoPascals (mB)
        /// </summary>
        internal static double Pressure
        {
            get
            {
                LogMessage("Pressure", (wx.Pres * 33.86).ToString());
                return wx.Pres * 33.86; // Weather file delivers inHg
            }
        }

        /// <summary>
        /// Rain rate at the observatory
        /// </summary>
        internal static double RainRate
        {
            get
            {
                LogMessage("RainRate", "get - not implemented");
                throw new PropertyNotImplementedException("RainRate", false);
            }
        }

        /// <summary>
        /// Forces the driver to immediately query its attached hardware to refresh sensor
        /// values
        /// </summary>
        internal static void Refresh()
        {
            throw new MethodNotImplementedException();
        }

        /// <summary>
        /// Provides a description of the sensor providing the requested property
        /// </summary>
        /// <param name="propertyName">Name of the property whose sensor description is required</param>
        /// <returns>The sensor description string</returns>
        internal static string SensorDescription(string propertyName)
        {
            switch (propertyName.Trim().ToLowerInvariant())
            {
                case "averageperiod":
                    return "Only immediate values are available.";
                case "cloudcover":
                case "dewpoint":
                case "humidity":
                case "pressure":
                case "skytemperature":
                case "temperature":
                case "winddirection":
                case "windgust":
                case "windspeed":
                case "skybrightness":
                    {

                        LogMessage("SensorDescription", propertyName + " is just this sensor, ya know?");
                        return propertyName + " is just this sensor, ya know?";
                    }

                // Throw an exception on the properties that are Not implemented
                case "starfwhm":
                case "rainrate":
                case "skyquality":
                    {
                        // Throw an exception on the properties that are not implemented
                        LogMessage("SensorDescription", $"Property {propertyName} is not implemented");
                        throw new MethodNotImplementedException($"SensorDescription - Property {propertyName} is not implemented");
                    }

                default:
                    {
                        LogMessage("SensorDescription", $"Invalid sensor name: {propertyName}");
                        throw new InvalidValueException($"SensorDescription - Invalid property name: {propertyName}");
                    }

            }
        }

        /// <summary>
        /// Sky brightness at the observatory
        /// </summary>
        internal static double SkyBrightness
        {
            get 
            {
                double value = 0.0;
                // A completely wild guess at lightsensor -> lux conversions, based on nothing more than a value of 673 in direct sunlight
                List<(int, int, double)> luxRanges = new List<(int, int, double)>
                {
                    (1024, 1021, .0001),
                    (1020, 1001, .002),
                    (1000, 981, 1.0),
                    (980, 951, 3.5),
                    (950, 751, 50),
                    (750, 746, 100),
                    (745, 741, 400),
                    (740, 721, 600),
                    (720, 651, 1000),
                    (650, 0, 32000.00)
                };
                foreach (var range in luxRanges)
                {
                    if (wx.LightSen <= range.Item1 && wx.LightSen > range.Item2)
                    {
                        value = range.Item3;
                        break;
                    }
                }
                LogMessage("SkyBrightness", value.ToString());
                return value;
            }
        }

        /// <summary>
        /// Sky quality at the observatory
        /// </summary>
        internal static double SkyQuality
        {
            get
            {
                LogMessage("SkyQuality", "get - not implemented");
                throw new PropertyNotImplementedException("SkyQuality", false);
            }
        }

        /// <summary>
        /// Seeing at the observatory
        /// </summary>
        internal static double StarFWHM
        {
            get
            {
                LogMessage("StarFWHM", "get - not implemented");
                throw new PropertyNotImplementedException("StarFWHM", false);
            }
        }

        /// <summary>
        /// Sky temperature at the observatory in deg C
        /// </summary>
        internal static double SkyTemperature
        {
            get
            {
                LogMessage("SkyTemperature", wx.SkyTemp.ToString());
                return wx.SkyTemp;
            }
        }

        /// <summary>
        /// Temperature at the observatory in deg C
        /// </summary>
        internal static double Temperature
        {
            get
            {
                LogMessage("Temperature", wx.Temp.ToString());
                return wx.Temp;
            }
        }

        /// <summary>
        /// Provides the time since the sensor value was last updated
        /// </summary>
        /// <param name="propertyName">Name of the property whose time since last update Is required</param>
        /// <returns>Time in seconds since the last sensor update for this property</returns>
        internal static double TimeSinceLastUpdate(string propertyName)
        {
            // Test for an empty property name, if found, return the time since the most recent update to any sensor
            double SecSinceLast;
            SecSinceLast = (DateTime.UtcNow - Convert.ToDateTime(wx.LastWrite)).TotalSeconds;
            if (!string.IsNullOrEmpty(propertyName))
            {
                switch (propertyName.ToLowerInvariant())
                {
                    // Return the time for properties that are implemented, otherwise fall through to the MethodNotImplementedException

                    case "cloudcover":
                    case "dewpoint":
                    case "humidity":
                    case "pressure":
                    case "skytemperature":
                    case "temperature":
                    case "winddirection":
                    case "windgust":
                    case "windspeed":
                    case "skybrightness":
                        {

                            LogMessage("TimeSinceLastUpdate", propertyName + " - " + SecSinceLast.ToString());
                            return SecSinceLast;
                        }

                    // Throw an exception on the properties that are Not implemented
                    case "starfwhm":
                    case "averageperiod":
                    case "rainrate":
                    case "skyquality":
                        {

                            LogMessage("TimeSinceLastUpdate", propertyName + " - Not Implemented");
                            // Unrecognized property name
                            throw new MethodNotImplementedException("TimeSinceLastUpdate(" + propertyName + ")");
                        }

                    default:
                        {
                            LogMessage("TimeSinceLastUpdate", propertyName + " - unrecognised");
                            throw new ASCOM.InvalidValueException("TimeSinceLastUpdate(" + propertyName + ")");
                        }
                }
            }
            else
            {
                // Return the time since the most recent update to any sensor
                LogMessage("TimeSinceLastUpdate", "No property - " + SecSinceLast.ToString());
                return SecSinceLast;
            }

        }

        /// <summary>
        /// Wind direction at the observatory in degrees
        /// </summary>
        internal static double WindDirection
        {
            get
            {
                LogMessage("WindDirection", wx.WDir.ToString());
                return wx.WDir;
            }
        }

        /// <summary>
        /// Peak 3 second wind gust at the observatory over the last 2 minutes in m/s
        /// </summary>
        internal static double WindGust
        {
            get
            {
                LogMessage("WindGust", wx.WGust.ToString());
                return wx.WGust;
            }
        }

        /// <summary>
        /// Wind speed at the observatory in m/s
        /// </summary>
        internal static double WindSpeed
        {
            get
            {
                LogMessage("WindSpeed", wx.WSp.ToString());
                return wx.WSp;
            }
        }

        #endregion

        #region Private methods

        #region Calculate the gust strength as the largest wind recorded over the last two minutes


        private static void UpdateGusts(double speed)
        {
            Dictionary<DateTime, double> newWinds = new Dictionary<DateTime, double>();
            var last = DateTime.Now - TimeSpan.FromMinutes(2);
            winds.Add(DateTime.Now, speed);
            var gust = 0.0;
            foreach (var item in winds)
            {
                if (item.Key > last)
                {
                    newWinds.Add(item.Key, item.Value);
                    if (item.Value > gust)
                        gust = item.Value;
                }
            }
            winds = newWinds;
        }

        #endregion

        #endregion

        #region Private properties and methods
        // Useful methods that can be used as required to help with driver development

        /// <summary>
        /// Returns true if there is a valid connection to the driver hardware
        /// </summary>
        private static bool IsConnected
        {
            get
            {
                // TODO check that the driver hardware connection exists and is connected to the hardware
                return connectedState;
            }
        }

        /// <summary>
        /// Use this function to throw an exception if we aren't connected to the hardware
        /// </summary>
        /// <param name="message"></param>
        private static void CheckConnected(string message)
        {
            if (!IsConnected)
            {
                throw new NotConnectedException(message);
            }
        }

        /// <summary>
        /// Read the device configuration from the ASCOM Profile store
        /// </summary>
        internal static void ReadProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "ObservingConditions";
                tl.Enabled = Convert.ToBoolean(driverProfile.GetValue(DriverProgId, traceStateProfileName, string.Empty, traceStateDefault));
                URL = driverProfile.GetValue(DriverProgId, URLProfileName, string.Empty, URLDefault);
            }
        }

        /// <summary>
        /// Write the device configuration to the  ASCOM  Profile store
        /// </summary>
        internal static void WriteProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "ObservingConditions";
                driverProfile.WriteValue(DriverProgId, traceStateProfileName, tl.Enabled.ToString());
                driverProfile.WriteValue(DriverProgId, URLProfileName, URL);
            }
        }

        /// <summary>
        /// Log helper function that takes identifier and message strings
        /// </summary>
        /// <param name="identifier"></param>
        /// <param name="message"></param>
        internal static void LogMessage(string identifier, string message)
        {
            tl.LogMessageCrLf(identifier, message);
        }

        /// <summary>
        /// Log helper function that takes formatted strings and arguments
        /// </summary>
        /// <param name="identifier"></param>
        /// <param name="message"></param>
        /// <param name="args"></param>
        internal static void LogMessage(string identifier, string message, params object[] args)
        {
            var msg = string.Format(message, args);
            LogMessage(identifier, msg);
        }

        private static void OnTimedEvent(object source, System.Timers.ElapsedEventArgs e)
        {
            wx.checkWeather(URL, tl);
        }

        #endregion
    }
}

