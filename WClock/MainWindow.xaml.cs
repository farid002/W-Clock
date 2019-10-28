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

    }
}
