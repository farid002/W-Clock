using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using Newtonsoft.Json;
using System.Windows.Media.Imaging;

namespace WClock
{
    class WeatherInfo
    {
        //Main class which will be used to call other classes
        public class root
        {
            public double latitude { get; set; }
            public double longitude { get; set; }
            public string timezone { get; set; }
            public daily daily { get; set; }
            public currently currently { get; set; }

        }

        // Data class which includes Daily weather characteristics
        public class data
        {
            public int time { get; set; }
            public string icon { get; set; }
            public float temperatureMin { get; set; }
            public float temperatureMax { get; set; }
            public int sunriseTime { get; set; }
            public int sunsetTime { get; set; }
            public float humidity { get; set; }
            public float windSpeed { get; set; }
            public string summary { get; set; }

        }

        // Daily class which includes Daily weather information as List
        public class daily
        {
            public List<data> data { get; set; }
        }

        // Currently class which includes Current weather information
        public class currently
        {
            public float temperature { get; set; }
            public string icon { get; set; }
            public float windSpeed { get; set; }
            public float humidity { get; set; }

        }

        public void getWeather()
        {
            //MainWindow mainWindow = new MainWindow();

            //new webClient
            using (WebClient web = new WebClient())
            {
                //Getting path for API
                string url = string.Format("https://api.darksky.net/forecast/c10099d11622db97e74edf8cbe651e7a/41.101417,29.029217?exclude=minutely,hourly,alerts&units=si&time=auto");
                var json = web.DownloadString(url);

                //Getting data from API as an object
                var result = JsonConvert.DeserializeObject<WeatherInfo.root>(json);

                WeatherInfo.root weatherData = result;
                string timezone = weatherData.timezone;
                double latitude = weatherData.latitude;
                double longitude = weatherData.longitude;

                
                DateTime day1 = new DateTime(1970, 1, 1, 0, 0, 0, 0);
                day1 = day1.AddSeconds(weatherData.daily.data[0].time).ToLocalTime();

                DateTime day2 = new DateTime(1970, 1, 1, 0, 0, 0, 0);
                day2 = day2.AddSeconds(weatherData.daily.data[1].time).ToLocalTime();

                DateTime day3 = new DateTime(1970, 1, 1, 0, 0, 0, 0);
                day3 = day3.AddSeconds(weatherData.daily.data[2].time).ToLocalTime();

                DateTime day4 = new DateTime(1970, 1, 1, 0, 0, 0, 0);
                day4 = day4.AddSeconds(weatherData.daily.data[3].time).ToLocalTime();

                DateTime day5 = new DateTime(1970, 1, 1, 0, 0, 0, 0);
                day5 = day5.AddSeconds(weatherData.daily.data[4].time).ToLocalTime();

                DateTime day6 = new DateTime(1970, 1, 1, 0, 0, 0, 0);
                day6 = day6.AddSeconds(weatherData.daily.data[5].time).ToLocalTime();

                DateTime day7 = new DateTime(1970, 1, 1, 0, 0, 0, 0);
                day7 = day7.AddSeconds(weatherData.daily.data[6].time).ToLocalTime();

                MainWindow.main.currWeather_Label.Content = weatherData.currently.temperature.ToString() + "'C";
                //MainWindow.main.CurrWeather_Image = new BitmapImage(new Uri(@"/Images/foo.png", UriKind.Relative));

                if (day1.DayOfWeek.ToString() == "Monday")
                {
                    /*mainWindow.MonWeather_Label.Content = "";
                    mainWindow.MonWeather_Image.Source = null; //new BitmapImage(new Uri(@"/Images/foo.png", UriKind.Relative));
                    mainWindow.currWeather_Label.Text = weatherData.currently.temperature.ToString() + "'C";*/


                }




            }
        }
    }

   
}
