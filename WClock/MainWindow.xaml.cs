using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Threading;
using System.Windows.Threading;
using System.Net;
using System.Web.Script.Serialization;
using outlook = Microsoft.Office.Interop.Outlook;

namespace WClock
{
    public partial class MainWindow : Window
    {
       
        public List<outlook.AppointmentItem> monList = new List<outlook.AppointmentItem>();
        public List<outlook.AppointmentItem> tueList = new List<outlook.AppointmentItem>();
        public List<outlook.AppointmentItem> wedList = new List<outlook.AppointmentItem>();
        public List<outlook.AppointmentItem> thuList = new List<outlook.AppointmentItem>();
        public List<outlook.AppointmentItem> friList = new List<outlook.AppointmentItem>();
        public List<outlook.AppointmentItem> satList = new List<outlook.AppointmentItem>();
        public List<outlook.AppointmentItem> sunList = new List<outlook.AppointmentItem>();

        public string location_string = "41.0082,28.9784"; //set location as Istanbul by default

        WeatherInfo weatherInfo = new WeatherInfo();
        MyCalendar outlookCalendar = new MyCalendar();

        int currentYear;
        int currentMonthDay;
        int currentHour;
        int currentMinute;
        int currentSec;
        string currentWeekDay;

        public static MainWindow main;

        public MainWindow()
        {
            InitializeComponent();
            StartClock();
            main = this;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            weatherInfo.getWeather();
            outlookCalendar.getAllAppointmentsForCurrentWeek();
        }

        //start the digital clock
        private void StartClock()
        {
            DispatcherTimer timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(1);
            timer.Tick += tickEvent;
            timer.Start();
        }

        //update timer box during each tick
        private void tickEvent(object sender, EventArgs e)
        {
            currentSec = DateTime.Now.Second;
            currentMinute = DateTime.Now.Minute;
            currentHour = DateTime.Now.Hour;
            currentWeekDay = DateTime.Now.DayOfWeek.ToString();
            currentMonthDay = DateTime.Now.Day;
            currentYear = DateTime.Now.Year;
            
            CurrentTime_Label.Content = currentHour.ToString("D2") + ":" + currentMinute.ToString("D2") + ":" + currentSec.ToString("D2");
            updateWeekTick();
        }

        //show day of week by the red arrow
        private void updateWeekTick()
        {
            mon_Line.Visibility = Visibility.Hidden;
            tue_Line.Visibility = Visibility.Hidden;
            wed_Line.Visibility = Visibility.Hidden;
            thu_Line.Visibility = Visibility.Hidden;
            fri_Line.Visibility = Visibility.Hidden;
            sat_Line.Visibility = Visibility.Hidden;
            sun_Line.Visibility = Visibility.Hidden;

            if (currentWeekDay == "Monday")
            {
                mon_Line.Visibility = Visibility.Visible;
            }
            else if(currentWeekDay == "Tuesday")
            {
                tue_Line.Visibility = Visibility.Visible;
            }
            else if (currentWeekDay == "Wednesday")
            {
                wed_Line.Visibility = Visibility.Visible;
            }
            else if (currentWeekDay == "Thursday")
            {
                thu_Line.Visibility = Visibility.Visible;
            }
            else if (currentWeekDay == "Friday")
            {
                fri_Line.Visibility = Visibility.Visible;
            }
            else if (currentWeekDay == "Saturday")
            {
                sat_Line.Visibility = Visibility.Visible;
            }
            else if (currentWeekDay == "Sunday")
            {
                sun_Line.Visibility = Visibility.Visible;
            }
        }

        public void MonWeather_Label_MouseEnter(object sender, MouseEventArgs e)
        {
            if(!DateTime.Now.DayOfWeek.Equals("Monday"))
                MonWeather_Label.ToolTip = "Minimum and Maximum degrees of the day in Celcius.\nWind speed:   ";
        }


        private void TueWeather_Label_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!DateTime.Now.DayOfWeek.Equals("Tuesday"))
                TueWeather_Label.ToolTip = "Minimum and Maximum degrees of the day in Celcius.";
        }

        private void WedWeather_Label_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!DateTime.Now.DayOfWeek.Equals("Wednesday"))
                WedWeather_Label.ToolTip = "Minimum and Maximum degrees of the day in Celcius.";
        }

        private void ThuWeather_Label_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!DateTime.Now.DayOfWeek.Equals("Thursday"))
                ThuWeather_Label.ToolTip = "Minimum and Maximum degrees of the day in Celcius.";
        }

        private void FriWeather_Label_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!DateTime.Now.DayOfWeek.Equals("Friday"))
                FriWeather_Label.ToolTip = "Minimum and Maximum degrees of the day in Celcius.";
        }

        private void SatWeather_Label_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!DateTime.Now.DayOfWeek.Equals("Saturday"))
                SatWeather_Label.ToolTip = "Minimum and Maximum degrees of the day in Celcius.";
        }

        private void SunWeather_Label_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!DateTime.Now.DayOfWeek.Equals("Sunday"))
                SunWeather_Label.ToolTip = "Minimum and Maximum degrees of the day in Celcius.";
        }

        private void Options_Button_Click(object sender, RoutedEventArgs e)
        {
            OptionsBorder.Visibility = Visibility.Visible;
            
        }
        private void Authors_Button_Click(object sender, RoutedEventArgs e)
        {
            OptionsBorder.Visibility = Visibility.Visible;

        }
        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        //Backgrounds
        //Blue
        private void Blue_Checked(object sender, RoutedEventArgs e)
        {
            Border_1.BorderBrush = new SolidColorBrush(Colors.Aqua);
            Border_2.BorderBrush = new SolidColorBrush(Colors.Aqua);
            Border_3.BorderBrush = new SolidColorBrush(Colors.Aqua);
            Border_4.BorderBrush = new SolidColorBrush(Colors.Aqua);
            Border_5.BorderBrush = new SolidColorBrush(Colors.Aqua);
            InsideLine_1.Stroke = new SolidColorBrush(Colors.Aqua);
            InsideLine_2.Stroke = new SolidColorBrush(Colors.Aqua);
            InsideLine_3.Stroke = new SolidColorBrush(Colors.Aqua);
            InsideLine_4.Stroke = new SolidColorBrush(Colors.Aqua);
            InsideLine_5.Stroke = new SolidColorBrush(Colors.Aqua);
            InsideLine_6.Stroke = new SolidColorBrush(Colors.Aqua);
            InsideLine_7.Stroke = new SolidColorBrush(Colors.Aqua);
            Border_2.Background = new SolidColorBrush(Colors.MidnightBlue);
            Border_3.Background = new SolidColorBrush(Colors.Blue);
            Border_4.Background = new SolidColorBrush(Colors.MidnightBlue);
            Border_5.Background = new SolidColorBrush(Colors.Blue);
            Mon_Label.Foreground = new SolidColorBrush(Colors.White);
            Tue_Label.Foreground = new SolidColorBrush(Colors.White);
            Wed_Label.Foreground = new SolidColorBrush(Colors.White);
            Thu_Label.Foreground = new SolidColorBrush(Colors.White);
            Fri_Label.Foreground = new SolidColorBrush(Colors.White);
            Sat_Label.Foreground = new SolidColorBrush(Colors.White);
            Sun_Label.Foreground = new SolidColorBrush(Colors.White);
            CurrentTime_Label.Foreground = new SolidColorBrush(Colors.White);
            currWeather_Label.Foreground = new SolidColorBrush(Colors.White);

        }
        // Red
        private void Red_Checked(object sender, RoutedEventArgs e)
        {
            Border_1.BorderBrush = new SolidColorBrush(Colors.Tomato);
            Border_2.BorderBrush = new SolidColorBrush(Colors.Tomato);
            Border_3.BorderBrush = new SolidColorBrush(Colors.Tomato);
            Border_4.BorderBrush = new SolidColorBrush(Colors.Tomato);
            Border_5.BorderBrush = new SolidColorBrush(Colors.Tomato);
            InsideLine_1.Stroke = new SolidColorBrush(Colors.Tomato);
            InsideLine_2.Stroke = new SolidColorBrush(Colors.Tomato);
            InsideLine_3.Stroke = new SolidColorBrush(Colors.Tomato);
            InsideLine_4.Stroke = new SolidColorBrush(Colors.Tomato);
            InsideLine_5.Stroke = new SolidColorBrush(Colors.Tomato);
            InsideLine_6.Stroke = new SolidColorBrush(Colors.Tomato);
            InsideLine_7.Stroke = new SolidColorBrush(Colors.Tomato);
            Border_2.Background = new SolidColorBrush(Colors.DarkRed);
            Border_3.Background = new SolidColorBrush(Colors.Red);
            Border_4.Background = new SolidColorBrush(Colors.DarkRed);
            Border_5.Background = new SolidColorBrush(Colors.Red);
            Mon_Label.Foreground = new SolidColorBrush(Colors.White);
            Tue_Label.Foreground = new SolidColorBrush(Colors.White);
            Wed_Label.Foreground = new SolidColorBrush(Colors.White);
            Thu_Label.Foreground = new SolidColorBrush(Colors.White);
            Fri_Label.Foreground = new SolidColorBrush(Colors.White);
            Sat_Label.Foreground = new SolidColorBrush(Colors.White);
            Sun_Label.Foreground = new SolidColorBrush(Colors.White);
            CurrentTime_Label.Foreground = new SolidColorBrush(Colors.White);
            currWeather_Label.Foreground = new SolidColorBrush(Colors.White);



        }
        // Green
        private void Green_Checked(object sender, RoutedEventArgs e)
        {
            Border_1.BorderBrush = new SolidColorBrush(Colors.LightGreen);
            Border_2.BorderBrush = new SolidColorBrush(Colors.LightGreen);
            Border_3.BorderBrush = new SolidColorBrush(Colors.LightGreen);
            Border_4.BorderBrush = new SolidColorBrush(Colors.LightGreen);
            Border_5.BorderBrush = new SolidColorBrush(Colors.LightGreen);
            InsideLine_1.Stroke = new SolidColorBrush(Colors.LightGreen);
            InsideLine_2.Stroke = new SolidColorBrush(Colors.LightGreen);
            InsideLine_3.Stroke = new SolidColorBrush(Colors.LightGreen);
            InsideLine_4.Stroke = new SolidColorBrush(Colors.LightGreen);
            InsideLine_5.Stroke = new SolidColorBrush(Colors.LightGreen);
            InsideLine_6.Stroke = new SolidColorBrush(Colors.LightGreen);
            InsideLine_7.Stroke = new SolidColorBrush(Colors.LightGreen);
            Border_2.Background = new SolidColorBrush(Colors.Green);
            Border_3.Background = new SolidColorBrush(Colors.Lime);
            Border_4.Background = new SolidColorBrush(Colors.Green);
            Border_5.Background = new SolidColorBrush(Colors.Lime);
            Mon_Label.Foreground = new SolidColorBrush(Colors.Black);
            Tue_Label.Foreground = new SolidColorBrush(Colors.Black);
            Wed_Label.Foreground = new SolidColorBrush(Colors.Black);
            Thu_Label.Foreground = new SolidColorBrush(Colors.Black);
            Fri_Label.Foreground = new SolidColorBrush(Colors.Black);
            Sat_Label.Foreground = new SolidColorBrush(Colors.Black);
            Sun_Label.Foreground = new SolidColorBrush(Colors.Black);
            CurrentTime_Label.Foreground = new SolidColorBrush(Colors.Black);
            currWeather_Label.Foreground = new SolidColorBrush(Colors.Black);
        }
        // Fenerbahce
        private void Fenerbahce_Checked(object sender, RoutedEventArgs e)
        {
            Border_1.BorderBrush = new SolidColorBrush(Colors.Yellow);
            Border_2.BorderBrush = new SolidColorBrush(Colors.Yellow);
            Border_3.BorderBrush = new SolidColorBrush(Colors.Yellow);
            Border_4.BorderBrush = new SolidColorBrush(Colors.Yellow);
            Border_5.BorderBrush = new SolidColorBrush(Colors.Yellow);
            InsideLine_1.Stroke = new SolidColorBrush(Colors.LightGreen);
            InsideLine_2.Stroke = new SolidColorBrush(Colors.LightGreen);
            InsideLine_3.Stroke = new SolidColorBrush(Colors.LightGreen);
            InsideLine_4.Stroke = new SolidColorBrush(Colors.LightGreen);
            InsideLine_5.Stroke = new SolidColorBrush(Colors.LightGreen);
            InsideLine_6.Stroke = new SolidColorBrush(Colors.LightGreen);
            InsideLine_7.Stroke = new SolidColorBrush(Colors.LightGreen);
            Border_2.Background = new SolidColorBrush(Colors.DarkBlue);
            Border_3.Background = new SolidColorBrush(Colors.Yellow);
            Border_4.Background = new SolidColorBrush(Colors.DarkBlue);
            Border_5.Background = new SolidColorBrush(Colors.Yellow);
            Mon_Label.Foreground = new SolidColorBrush(Colors.Black);
            Tue_Label.Foreground = new SolidColorBrush(Colors.Black);
            Wed_Label.Foreground = new SolidColorBrush(Colors.Black);
            Thu_Label.Foreground = new SolidColorBrush(Colors.Black);
            Fri_Label.Foreground = new SolidColorBrush(Colors.Black);
            Sat_Label.Foreground = new SolidColorBrush(Colors.Black);
            Sun_Label.Foreground = new SolidColorBrush(Colors.Black);
            CurrentTime_Label.Foreground = new SolidColorBrush(Colors.Black);
            currWeather_Label.Foreground = new SolidColorBrush(Colors.Black);
        }

        // Locations
        // Baku
        private void Baku_Checked(object sender, RoutedEventArgs e)
        {
            location_string = "40.4093,49.8671";
        }
        // Berlin
        private void Berlin_Checked(object sender, RoutedEventArgs e)
        {
            location_string = "52.5200,13.4050";
        }
        // Istanbul
        private void Istanbul_Checked(object sender, RoutedEventArgs e)
        {
            location_string = "41.0082,28.9784";
        }
        // London
        private void London_Checked(object sender, RoutedEventArgs e)
        {
            location_string = "51.5074,0.1278";
        }
        //Moscow
        private void Moscow_Checked(object sender, RoutedEventArgs e)
        {
            location_string = "55.7558,37.6173";
        }
        // NewYork
        private void NewYork_Checked(object sender, RoutedEventArgs e)
        {
            location_string = "40.730610,-73.935242";
        }
        //Paris
        private void Paris_Checked(object sender, RoutedEventArgs e)
        {
            location_string = "48.8566,2.3522";
        }

        // Apply the new Location.
        private void LocationApply_Button_Click(object sender, RoutedEventArgs e)
        {
            weatherInfo.getWeather();
        }

        // Apply new background
        private void BackgroundApply_Button_Click(object sender, RoutedEventArgs e)
        {
            weatherInfo.getWeather();
        }

        // Fonts
        // Calibri
        private void Calibri_Checked(object sender, RoutedEventArgs e)
        {
            Application.Current.MainWindow.FontFamily = new FontFamily("Calibri");
        }
        // Century Gothic 
        private void Century_Gothic_Checked(object sender, RoutedEventArgs e)
        {
            Application.Current.MainWindow.FontFamily = new FontFamily("Century Gothic");
        }
        // Goudy Old Style
        private void Goudy_Old_Style_Checked(object sender, RoutedEventArgs e)
        {
            Application.Current.MainWindow.FontFamily = new FontFamily("Goudy Old Style");
        }
        // Brush Script MT
        private void Brush_Script_MT_Checked(object sender, RoutedEventArgs e)
        {
            Application.Current.MainWindow.FontFamily = new FontFamily("Brush Script MT");
        }
        // Papyrus
        private void Papyrus_Checked(object sender, RoutedEventArgs e)
        {
            Application.Current.MainWindow.FontFamily = new FontFamily("papyrus");
        }
        //Ink Free
        private void Ink_Free_Checked(object sender, RoutedEventArgs e)
        {
            Application.Current.MainWindow.FontFamily = new FontFamily("Ink Free");
        }
        //Default
        private void Default_Checked(object sender, RoutedEventArgs e)
        {
            Application.Current.MainWindow.FontFamily = new FontFamily("Default");
        }

        private void FontApply_Button_Click(object sender, RoutedEventArgs e)
        {

        }

        // Close Button
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }
    }
}
