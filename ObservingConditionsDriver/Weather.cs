using System;
using ASCOM.Utilities;
using Newtonsoft.Json;

namespace ASCOM.LocalServer
{
    class Weather
    {
        public string LastWrite { get; set; }
        public long LastWrite_timestamp { get; set; }
        public double Temp { get; set; }
        public double Hum { get; set; }
        public double DewPoint { get; set; }
        public double Pres { get; set; }
        public double WSp { get; set; }
        public double WGust { get; set; }
        public double WDir { get; set; }
        public double RTot { get; set; }
        public int LightSen { get; set; }
        public double SkyTemp { get; set; }
        public double IRAmb { get; set; }
        public int RSen { get; set; }
        public int CloudCondition { get; set; }
        public int DaylightCondition { get; set; }
        public int RainCondition { get; set; }
        public int WindCondition { get; set; }
        public int RSenD { get; set; }
        public string UnsafeWarning { get; set; }
        public int Alert { get; set; }

        public double MPHtoMPS(double MPH)
        {
            double MPHtoMPSRet = default;
            MPHtoMPSRet = MPH / 2.237d;
            return MPHtoMPSRet;
        }
        public void checkWeather(string URL, TraceLogger logger)
        {
            Weather updatedWeather = new Weather();
            try
            {
                // wx = New Weather
                var webclient = new Utils.WebClient();
                string response = webclient.DownloadString(URL);
                updatedWeather = JsonConvert.DeserializeObject<LocalServer.Weather>(response);
                logger.LogMessage("checkWeather", "Read JSON at " + DateTime.Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss") + " UTC");
                logger.LogMessage("checkWeather", "LastWrite at " + Convert.ToDateTime(updatedWeather.LastWrite).ToString("yyyy-MM-dd HH:mm:ss") + " UTC");
                this.LastWrite = updatedWeather.LastWrite;
                this.LastWrite_timestamp = updatedWeather.LastWrite_timestamp;
                this.Temp = updatedWeather.Temp;
                this.Hum = updatedWeather.Hum;
                this.Pres = updatedWeather.Pres;
                this.WDir = updatedWeather.WDir;
                this.WSp = updatedWeather.WSp;
                this.WGust = updatedWeather.WGust;
                this.WDir = updatedWeather.WDir;
                this.RTot = updatedWeather.RTot;
                this.LightSen = updatedWeather.LightSen;
                this.SkyTemp = updatedWeather.SkyTemp;
                this.IRAmb = updatedWeather.IRAmb;
                this.CloudCondition = updatedWeather.CloudCondition;
                this.DaylightCondition = updatedWeather.DaylightCondition;
                this.RainCondition = updatedWeather.RainCondition;
                this.WindCondition = updatedWeather.WindCondition;
                this.RSenD = updatedWeather.RSenD;
                this.UnsafeWarning = updatedWeather.UnsafeWarning;
                this.Alert = updatedWeather.Alert;
                this.DewPoint = updatedWeather.DewPoint;
            }
            catch (Exception ex)
            {
                    logger.LogMessage("checkWeather ERROR", ex.Message);
            }
        }
    }
}
