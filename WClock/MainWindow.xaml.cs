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
        List<outlook.AppointmentItem> monList = new List<outlook.AppointmentItem>();
        List<outlook.AppointmentItem> tueList = new List<outlook.AppointmentItem>();
        List<outlook.AppointmentItem> wedList = new List<outlook.AppointmentItem>();
        List<outlook.AppointmentItem> thuList = new List<outlook.AppointmentItem>();
        List<outlook.AppointmentItem> friList = new List<outlook.AppointmentItem>();
        List<outlook.AppointmentItem> satList = new List<outlook.AppointmentItem>();
        List<outlook.AppointmentItem> sunList = new List<outlook.AppointmentItem>();

        string sunnyWeatherImagePath = "imgs/weather-icons/sunny.png";
        string cloudyWeatherImagePath = "imgs/weather-icons/cloudy.png";
        string cloudysunnyWeatherImagePath = "imgs/weather-icons/cloudysunny.png";
        string snowyWeatherImagePath = "imgs/weather-icons/snowy.png";
        string lowrainyWeatherImagePath = "imgs/weather-icons/lowrainy.png";
        string rainyWeatherImagePath = "imgs/weather-icons/rainy.png";
        string lightningWeatherImagePath = "imgs/weather-icons/lighning.png";
        string lightningrainyWeatherImagePath = "imgs/weather-icons/lighningrainy.png";

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
            WeatherInfo info = new WeatherInfo();
            info.getWeather();
            
            getAllAppointmentsForCurrentWeek();
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

        //create an event/appointment
        private void createAppointment()
        {
            outlook.Application app = new outlook.Application();
            outlook.AppointmentItem appoinment = (outlook.AppointmentItem)app.CreateItem(outlook.OlItemType.olAppointmentItem);
            appoinment.Body = "Somethingg";
            appoinment.Importance = outlook.OlImportance.olImportanceNormal;
            appoinment.Save();   //  =  ((outlook._AppointmentItem)appoinment).Save();
            MessageBox.Show(appoinment.UserProperties.ToString());
        }

        //get all your appointments/events from ms outlook for the next 7 days
        private void getAllAppointmentsForCurrentWeek()
        {
            DateTime currDateTime = DateTime.Now.Date;

            outlook._Application app = new outlook.Application();
            outlook._NameSpace ns = app.GetNamespace("MAPI");
            outlook.MAPIFolder mAPIFolder = ns.GetDefaultFolder(outlook.OlDefaultFolders.olFolderCalendar);
            foreach (outlook.AppointmentItem item in mAPIFolder.Items)
            {
                if(item.Start.Day == currDateTime.Day)
                {
                    if(item.Start.DayOfWeek.ToString() == "Monday")
                    {
                        monList.Add(item);
                        if (monApp_listbox.Items.Count < 2)
                            monApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if(item.Start.DayOfWeek.ToString() == "Tuesday")
                    {
                        tueList.Add(item);
                        if (tueApp_listbox.Items.Count < 2)
                            tueApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Wednesday")
                    {
                        wedList.Add(item);
                        if (wedApp_listbox.Items.Count < 2)
                            wedApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Thursday")
                    {
                        thuList.Add(item);
                        if (thuApp_listbox.Items.Count < 2)
                            thuApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Friday")
                    {
                        friList.Add(item);
                        if (friApp_listbox.Items.Count < 2)
                            friApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Saturday")
                    {
                        satList.Add(item);
                        if (satApp_listbox.Items.Count < 2)
                            satApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Sunday")
                    {
                        sunList.Add(item);
                        if (sunApp_listbox.Items.Count < 2)
                            sunApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                }
                else if (item.Start.Day == currDateTime.AddDays(1).Day)
                {
                    if (item.Start.DayOfWeek.ToString() == "Monday")
                    {
                        monList.Add(item);
                        if (monApp_listbox.Items.Count < 2)
                            monApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Tuesday")
                    {
                        tueList.Add(item);
                        if (tueApp_listbox.Items.Count < 2)
                            tueApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Wednesday")
                    {
                        wedList.Add(item);
                        if (wedApp_listbox.Items.Count < 2)
                            wedApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Thursday")
                    {
                        thuList.Add(item);
                        if (thuApp_listbox.Items.Count < 2)
                            thuApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Friday")
                    {
                        friList.Add(item);
                        if (friApp_listbox.Items.Count < 2)
                            friApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Saturday")
                    {
                        satList.Add(item);
                        if (satApp_listbox.Items.Count < 2)
                            satApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Sunday")
                    {
                        sunList.Add(item);
                        if (sunApp_listbox.Items.Count < 2)
                            sunApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                }
                else if (item.Start.Day == currDateTime.AddDays(2).Day)
                {
                    if (item.Start.DayOfWeek.ToString() == "Monday")
                    {
                        monList.Add(item);
                        if (monApp_listbox.Items.Count < 2)
                            monApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Tuesday")
                    {
                        tueList.Add(item);
                        if (tueApp_listbox.Items.Count < 2)
                            tueApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Wednesday")
                    {
                        wedList.Add(item);
                        if (wedApp_listbox.Items.Count < 2)
                            wedApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Thursday")
                    {
                        thuList.Add(item);
                        if (thuApp_listbox.Items.Count < 2)
                            thuApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Friday")
                    {
                        friList.Add(item);
                        if (friApp_listbox.Items.Count < 2)
                            friApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Saturday")
                    {
                        satList.Add(item);
                        if (satApp_listbox.Items.Count < 2)
                            satApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Sunday")
                    {
                        sunList.Add(item);
                        if (sunApp_listbox.Items.Count < 2)
                            sunApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                }
                else if (item.Start.Day == currDateTime.AddDays(3).Day)
                {
                    if (item.Start.DayOfWeek.ToString() == "Monday")
                    {
                        monList.Add(item);
                        if (monApp_listbox.Items.Count < 2)
                            monApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Tuesday")
                    {
                        tueList.Add(item);
                        if (tueApp_listbox.Items.Count < 2)
                            tueApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Wednesday")
                    {
                        wedList.Add(item);
                        if (wedApp_listbox.Items.Count < 2)
                            wedApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Thursday")
                    {
                        thuList.Add(item);
                        if (thuApp_listbox.Items.Count < 2)
                            thuApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Friday")
                    {
                        friList.Add(item);
                        if (friApp_listbox.Items.Count < 2)
                            friApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Saturday")
                    {
                        satList.Add(item);
                        if (satApp_listbox.Items.Count < 2)
                            satApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Sunday")
                    {
                        sunList.Add(item);
                        if (sunApp_listbox.Items.Count < 2)
                            sunApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                }
                else if (item.Start.Day == currDateTime.AddDays(4).Day)
                {
                    if (item.Start.DayOfWeek.ToString() == "Monday")
                    {
                        monList.Add(item);
                        if (monApp_listbox.Items.Count < 2)
                            monApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Tuesday")
                    {
                        tueList.Add(item);
                        if (tueApp_listbox.Items.Count < 2)
                            tueApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Wednesday")
                    {
                        wedList.Add(item);
                        if (wedApp_listbox.Items.Count < 2)
                            wedApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Thursday")
                    {
                        thuList.Add(item);
                        if (thuApp_listbox.Items.Count < 2)
                            thuApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Friday")
                    {
                        friList.Add(item);
                        if (friApp_listbox.Items.Count < 2)
                            friApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Saturday")
                    {
                        satList.Add(item);
                        if (satApp_listbox.Items.Count < 2)
                            satApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Sunday")
                    {
                        sunList.Add(item);
                        if (sunApp_listbox.Items.Count < 2)
                            sunApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                }
                else if (item.Start.Day == currDateTime.AddDays(5).Day)
                {
                    if (item.Start.DayOfWeek.ToString() == "Monday")
                    {
                        monList.Add(item);
                        if (monApp_listbox.Items.Count < 2)
                            monApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Tuesday")
                    {
                        tueList.Add(item);
                        if (tueApp_listbox.Items.Count < 2)
                            tueApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Wednesday")
                    {
                        wedList.Add(item);
                        if (wedApp_listbox.Items.Count < 2)
                            wedApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Thursday")
                    {
                        thuList.Add(item);
                        if (thuApp_listbox.Items.Count < 2)
                            thuApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Friday")
                    {
                        friList.Add(item);
                        if (friApp_listbox.Items.Count < 2)
                            friApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Saturday")
                    {
                        satList.Add(item);
                        if (satApp_listbox.Items.Count < 2)
                            satApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Sunday")
                    {
                        sunList.Add(item);
                        if (sunApp_listbox.Items.Count < 2)
                            sunApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                }
                else if (item.Start.Day == currDateTime.AddDays(6).Day)
                {
                    if (item.Start.DayOfWeek.ToString() == "Monday")
                    {
                        monList.Add(item);
                        if (monApp_listbox.Items.Count < 2)
                            monApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Tuesday")
                    {
                        tueList.Add(item);
                        if (tueApp_listbox.Items.Count < 2)
                            tueApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Wednesday")
                    {
                        wedList.Add(item);
                        if (wedApp_listbox.Items.Count < 2)
                            wedApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Thursday")
                    {
                        thuList.Add(item);
                        if (thuApp_listbox.Items.Count < 2)
                            thuApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Friday")
                    {
                        friList.Add(item);
                        if (friApp_listbox.Items.Count < 2)
                            friApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Saturday")
                    {
                        satList.Add(item);
                        if (satApp_listbox.Items.Count < 2)
                            satApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Sunday")
                    {
                        sunList.Add(item);
                        if (sunApp_listbox.Items.Count < 2)
                            sunApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                }
            }
            if (monList.Count >= 3) monApp_listbox.Items.Add("...more...");
            if (tueList.Count >= 3) tueApp_listbox.Items.Add("...more...");
            if (wedList.Count >= 3) wedApp_listbox.Items.Add("...more...");
            if (thuList.Count >= 3) thuApp_listbox.Items.Add("...more...");
            if (friList.Count >= 3) friApp_listbox.Items.Add("...more...");
            if (satList.Count >= 3) satApp_listbox.Items.Add("...more...");
            if (sunList.Count >= 3) sunApp_listbox.Items.Add("...more...");
        }

        private void OptionButton_Click(object sender, RoutedEventArgs e)
        {
            OptionsBorder.Visibility = Visibility.Visible;
            
        }

        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void Blue_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void Red_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void Green_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void Fenerbahce_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            
        }
    }
}
