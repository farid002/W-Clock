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
            string weatherImagePath = "/imgs/weather-icons/";

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

                MainWindow.main.currWeather_Label.Content = ((int)weatherData.currently.temperature).ToString() + "°C";
                MainWindow.main.CurrWeather_Image.Source = new BitmapImage(new Uri( weatherImagePath + weatherData.currently.icon + @".png", UriKind.Relative));

                if (day1.DayOfWeek.ToString() == "Monday") //if today is monday
                {
                    //day2 is Tue
                    MainWindow.main.TueWeather_Label.Content = ((int)weatherData.daily.data[1].temperatureMin).ToString() + "° / " 
                                                             + ((int)weatherData.daily.data[1].temperatureMax).ToString() + "°";
                    MainWindow.main.TueWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[1].icon + @".png", UriKind.Relative));
                    
                    //day3 is Wed
                    MainWindow.main.WedWeather_Label.Content = ((int)weatherData.daily.data[2].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[2].temperatureMax).ToString() + "°";
                    MainWindow.main.WedWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[2].icon + @".png", UriKind.Relative));

                    //day4 is Thu
                    MainWindow.main.ThuWeather_Label.Content = ((int)weatherData.daily.data[3].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[3].temperatureMax).ToString() + "°";
                    MainWindow.main.ThuWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[3].icon + @".png", UriKind.Relative));

                    //day5 is Fri
                    MainWindow.main.FriWeather_Label.Content = ((int)weatherData.daily.data[4].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[4].temperatureMax).ToString() + "°";
                    MainWindow.main.FriWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[4].icon + @".png", UriKind.Relative));

                    //day6 is Sat
                    MainWindow.main.SatWeather_Label.Content = ((int)weatherData.daily.data[5].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[5].temperatureMax).ToString() + "°";
                    MainWindow.main.SatWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[5].icon + @".png", UriKind.Relative));

                    //day7 is Sun
                    MainWindow.main.SunWeather_Label.Content = ((int)weatherData.daily.data[6].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[6].temperatureMax).ToString() + "°";
                    MainWindow.main.SunWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[6].icon + @".png", UriKind.Relative));

                }

                else if (day1.DayOfWeek.ToString() == "Tuesday") //if today is tuesday
                {
                    //day7 is Mon
                    MainWindow.main.MonWeather_Label.Content = ((int)weatherData.daily.data[6].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[6].temperatureMax).ToString() + "°";
                    MainWindow.main.MonWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[6].icon + @".png", UriKind.Relative));

                    //day1  is Tue
                    MainWindow.main.TueWeather_Label.Content = "";
                    MainWindow.main.TueWeather_Image.Source = null;

                    //day2 is Wed
                    MainWindow.main.WedWeather_Label.Content = ((int)weatherData.daily.data[1].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[1].temperatureMax).ToString() + "°";
                    MainWindow.main.WedWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[1].icon + @".png", UriKind.Relative));

                    //day3 is Thu
                    MainWindow.main.ThuWeather_Label.Content = ((int)weatherData.daily.data[2].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[2].temperatureMax).ToString() + "°";
                    MainWindow.main.ThuWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[2].icon + @".png", UriKind.Relative));

                    //day4 is Fri
                    MainWindow.main.FriWeather_Label.Content = ((int)weatherData.daily.data[3].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[3].temperatureMax).ToString() + "°";
                    MainWindow.main.FriWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[3].icon + @".png", UriKind.Relative));

                    //day5 is Sat
                    MainWindow.main.SatWeather_Label.Content = ((int)weatherData.daily.data[4].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[4].temperatureMax).ToString() + "°";
                    MainWindow.main.SatWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[4].icon + @".png", UriKind.Relative));

                    //day6 is Sun
                    MainWindow.main.SunWeather_Label.Content = ((int)weatherData.daily.data[5].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[5].temperatureMax).ToString() + "°";
                    MainWindow.main.SunWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[5].icon + @".png", UriKind.Relative));


                }
                else if (day1.DayOfWeek.ToString() == "Wednesday") //if today is wed
                {
                    //day6 is Mon
                    MainWindow.main.MonWeather_Label.Content = ((int)weatherData.daily.data[5].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[5].temperatureMax).ToString() + "°";
                    MainWindow.main.MonWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[5].icon + @".png", UriKind.Relative));

                    //day7 is Tue
                    MainWindow.main.TueWeather_Label.Content = ((int)weatherData.daily.data[6].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[6].temperatureMax).ToString() + "°";
                    MainWindow.main.TueWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[6].icon + @".png", UriKind.Relative));

                    //day1 is Wed
                    MainWindow.main.WedWeather_Label.Content = "";
                    MainWindow.main.WedWeather_Image.Source = null;

                    //day2 is Thu
                    MainWindow.main.ThuWeather_Label.Content = ((int)weatherData.daily.data[1].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[1].temperatureMax).ToString() + "°";
                    MainWindow.main.ThuWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[1].icon + @".png", UriKind.Relative));

                    //day3 is Fri
                    MainWindow.main.FriWeather_Label.Content = ((int)weatherData.daily.data[2].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[2].temperatureMax).ToString() + "°";
                    MainWindow.main.FriWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[2].icon + @".png", UriKind.Relative));

                    //day4 is Sat
                    MainWindow.main.SatWeather_Label.Content = ((int)weatherData.daily.data[3].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[3].temperatureMax).ToString() + "°";
                    MainWindow.main.SatWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[3].icon + @".png", UriKind.Relative));

                    //day5 is Sun
                    MainWindow.main.SunWeather_Label.Content = ((int)weatherData.daily.data[4].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[4].temperatureMax).ToString() + "°";
                    MainWindow.main.SunWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[4].icon + @".png", UriKind.Relative));


                }
                else if (day1.DayOfWeek.ToString() == "Thursday") //if today is thu
                {
                    //day5 is Mon
                    MainWindow.main.MonWeather_Label.Content = ((int)weatherData.daily.data[4].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[4].temperatureMax).ToString() + "°";
                    MainWindow.main.MonWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[4].icon + @".png", UriKind.Relative));

                    //day6 is Tue
                    MainWindow.main.TueWeather_Label.Content = ((int)weatherData.daily.data[5].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[5].temperatureMax).ToString() + "°";
                    MainWindow.main.TueWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[5].icon + @".png", UriKind.Relative));

                    //day7 is Wed
                    MainWindow.main.WedWeather_Label.Content = ((int)weatherData.daily.data[6].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[6].temperatureMax).ToString() + "°";
                    MainWindow.main.WedWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[6].icon + @".png", UriKind.Relative));

                    //day1 is Thu
                    MainWindow.main.ThuWeather_Label.Content = "";
                    MainWindow.main.ThuWeather_Image.Source = null;

                    //day2 is Fri
                    MainWindow.main.FriWeather_Label.Content = ((int)weatherData.daily.data[1].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[1].temperatureMax).ToString() + "°";
                    MainWindow.main.FriWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[1].icon + @".png", UriKind.Relative));

                    //day3 is Sat
                    MainWindow.main.SatWeather_Label.Content = ((int)weatherData.daily.data[2].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[2].temperatureMax).ToString() + "°";
                    MainWindow.main.SatWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[2].icon + @".png", UriKind.Relative));

                    //day4 is Sun
                    MainWindow.main.SunWeather_Label.Content = ((int)weatherData.daily.data[3].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[3].temperatureMax).ToString() + "°";
                    MainWindow.main.SunWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[3].icon + @".png", UriKind.Relative));


                }
                else if (day1.DayOfWeek.ToString() == "Friday") //if today is fri
                {
                    //day4 is Mon
                    MainWindow.main.MonWeather_Label.Content = ((int)weatherData.daily.data[3].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[3].temperatureMax).ToString() + "°";
                    MainWindow.main.MonWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[3].icon + @".png", UriKind.Relative));

                    //day5 is Tue
                    MainWindow.main.TueWeather_Label.Content = ((int)weatherData.daily.data[4].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[4].temperatureMax).ToString() + "°";
                    MainWindow.main.TueWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[4].icon + @".png", UriKind.Relative));

                    //day6 is Wed
                    MainWindow.main.WedWeather_Label.Content = ((int)weatherData.daily.data[5].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[5].temperatureMax).ToString() + "°";
                    MainWindow.main.WedWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[5].icon + @".png", UriKind.Relative));

                    //day7 is Thu
                    MainWindow.main.ThuWeather_Label.Content = ((int)weatherData.daily.data[6].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[6].temperatureMax).ToString() + "°";
                    MainWindow.main.ThuWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[6].icon + @".png", UriKind.Relative));

                    //day1 is Fri
                    MainWindow.main.FriWeather_Label.Content = "";
                    MainWindow.main.FriWeather_Image.Source = null;

                    //day2 is Sat
                    MainWindow.main.SatWeather_Label.Content = ((int)weatherData.daily.data[1].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[1].temperatureMax).ToString() + "°";
                    MainWindow.main.SatWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[1].icon + @".png", UriKind.Relative));

                    //day3 is Sun
                    MainWindow.main.SunWeather_Label.Content = ((int)weatherData.daily.data[2].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[2].temperatureMax).ToString() + "°";
                    MainWindow.main.SunWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[2].icon + @".png", UriKind.Relative));
                }

                else if (day1.DayOfWeek.ToString() == "Saturday") //if today is sat
                {
                    //day3 is Mon
                    MainWindow.main.MonWeather_Label.Content = ((int)weatherData.daily.data[2].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[2].temperatureMax).ToString() + "°";
                    MainWindow.main.MonWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[2].icon + @".png", UriKind.Relative));

                    //day4 is Tue
                    MainWindow.main.TueWeather_Label.Content = ((int)weatherData.daily.data[3].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[3].temperatureMax).ToString() + "°";
                    MainWindow.main.TueWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[3].icon + @".png", UriKind.Relative));

                    //day5 is Wed
                    MainWindow.main.WedWeather_Label.Content = ((int)weatherData.daily.data[4].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[4].temperatureMax).ToString() + "°";
                    MainWindow.main.WedWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[4].icon + @".png", UriKind.Relative));

                    //day6 is Thu
                    MainWindow.main.ThuWeather_Label.Content = ((int)weatherData.daily.data[5].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[5].temperatureMax).ToString() + "°";
                    MainWindow.main.ThuWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[5].icon + @".png", UriKind.Relative));

                    //day7 is Fri
                    MainWindow.main.FriWeather_Label.Content = ((int)weatherData.daily.data[6].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[6].temperatureMax).ToString() + "°";
                    MainWindow.main.FriWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[6].icon + @".png", UriKind.Relative));

                    //day1 is Sat
                    MainWindow.main.SatWeather_Label.Content = "";
                    MainWindow.main.SatWeather_Image.Source = null;

                    //day2 is Sun
                    MainWindow.main.SunWeather_Label.Content = ((int)weatherData.daily.data[1].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[1].temperatureMax).ToString() + "°";
                    MainWindow.main.SunWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[1].icon + @".png", UriKind.Relative));
                }
                else if (day1.DayOfWeek.ToString() == "Sunday") //if today is Sun
                {
                    //day2 is Mon
                    MainWindow.main.MonWeather_Label.Content = ((int)weatherData.daily.data[1].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[1].temperatureMax).ToString() + "°";
                    MainWindow.main.MonWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[1].icon + @".png", UriKind.Relative));

                    //day3 is Tue
                    MainWindow.main.TueWeather_Label.Content = ((int)weatherData.daily.data[2].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[2].temperatureMax).ToString() + "°";
                    MainWindow.main.TueWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[2].icon + @".png", UriKind.Relative));

                    //day4 is Wed
                    MainWindow.main.WedWeather_Label.Content = ((int)weatherData.daily.data[3].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[3].temperatureMax).ToString() + "°";
                    MainWindow.main.WedWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[3].icon + @".png", UriKind.Relative));

                    //day5 is Thu
                    MainWindow.main.ThuWeather_Label.Content = ((int)weatherData.daily.data[4].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[4].temperatureMax).ToString() + "°";
                    MainWindow.main.ThuWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[4].icon + @".png", UriKind.Relative));

                    //day6 is Fri
                    MainWindow.main.FriWeather_Label.Content = ((int)weatherData.daily.data[5].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[5].temperatureMax).ToString() + "°";
                    MainWindow.main.FriWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[5].icon + @".png", UriKind.Relative));

                    //day7 is Sat
                    MainWindow.main.SatWeather_Label.Content = ((int)weatherData.daily.data[6].temperatureMin).ToString() + "° / "
                                                             + ((int)weatherData.daily.data[6].temperatureMax).ToString() + "°";
                    MainWindow.main.SatWeather_Image.Source = new BitmapImage(new Uri(weatherImagePath + weatherData.daily.data[6].icon + @".png", UriKind.Relative));

                    //day1 is Sun
                    MainWindow.main.SunWeather_Label.Content = "";
                    MainWindow.main.SunWeather_Image.Source = null;
                }
            }
        }
    }

   
}
