using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using outlook = Microsoft.Office.Interop.Outlook;

namespace WClock
{
    class MyCalendar
    {
        public void getAllAppointmentsForCurrentWeek()
        {
            DateTime currDateTime = DateTime.Now.Date;

            outlook._Application app = new outlook.Application();
            outlook._NameSpace ns = app.GetNamespace("MAPI");
            outlook.MAPIFolder mAPIFolder = ns.GetDefaultFolder(outlook.OlDefaultFolders.olFolderCalendar);

            //add items into lists and listbox elements
            foreach (outlook.AppointmentItem item in mAPIFolder.Items)
            {
                if (item.Start.Day == currDateTime.Day)
                {
                    if (item.Start.DayOfWeek.ToString() == "Monday")
                    {
                        MainWindow.main.monList.Add(item);
                        if (MainWindow.main.monApp_listbox.Items.Count < 2)
                            MainWindow.main.monApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Tuesday")
                    {
                        MainWindow.main.tueList.Add(item);
                        if (MainWindow.main.tueApp_listbox.Items.Count < 2)
                            MainWindow.main.tueApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Wednesday")
                    {
                        MainWindow.main.wedList.Add(item);
                        if (MainWindow.main.wedApp_listbox.Items.Count < 2)
                            MainWindow.main.wedApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Thursday")
                    {
                        MainWindow.main.thuList.Add(item);
                        if (MainWindow.main.thuApp_listbox.Items.Count < 2)
                            MainWindow.main.thuApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Friday")
                    {
                        MainWindow.main.friList.Add(item);
                        if (MainWindow.main.friApp_listbox.Items.Count < 2)
                            MainWindow.main.friApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Saturday")
                    {
                        MainWindow.main.satList.Add(item);
                        if (MainWindow.main.satApp_listbox.Items.Count < 2)
                            MainWindow.main.satApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Sunday")
                    {
                        MainWindow.main.sunList.Add(item);
                        if (MainWindow.main.sunApp_listbox.Items.Count < 2)
                            MainWindow.main.sunApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                }
                else if (item.Start.Day == currDateTime.AddDays(1).Day)
                {
                    if (item.Start.DayOfWeek.ToString() == "Monday")
                    {
                        MainWindow.main.monList.Add(item);
                        if (MainWindow.main.monApp_listbox.Items.Count < 2)
                            MainWindow.main.monApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Tuesday")
                    {
                        MainWindow.main.tueList.Add(item);
                        if (MainWindow.main.tueApp_listbox.Items.Count < 2)
                            MainWindow.main.tueApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Wednesday")
                    {
                        MainWindow.main.wedList.Add(item);
                        if (MainWindow.main.wedApp_listbox.Items.Count < 2)
                            MainWindow.main.wedApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Thursday")
                    {
                        MainWindow.main.thuList.Add(item);
                        if (MainWindow.main.thuApp_listbox.Items.Count < 2)
                            MainWindow.main.thuApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Friday")
                    {
                        MainWindow.main.friList.Add(item);
                        if (MainWindow.main.friApp_listbox.Items.Count < 2)
                            MainWindow.main.friApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Saturday")
                    {
                        MainWindow.main.satList.Add(item);
                        if (MainWindow.main.satApp_listbox.Items.Count < 2)
                            MainWindow.main.satApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Sunday")
                    {
                        MainWindow.main.sunList.Add(item);
                        if (MainWindow.main.sunApp_listbox.Items.Count < 2)
                            MainWindow.main.sunApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                }
                else if (item.Start.Day == currDateTime.AddDays(2).Day)
                {
                    if (item.Start.DayOfWeek.ToString() == "Monday")
                    {
                        MainWindow.main.monList.Add(item);
                        if (MainWindow.main.monApp_listbox.Items.Count < 2)
                            MainWindow.main.monApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Tuesday")
                    {
                        MainWindow.main.tueList.Add(item);
                        if (MainWindow.main.tueApp_listbox.Items.Count < 2)
                            MainWindow.main.tueApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Wednesday")
                    {
                        MainWindow.main.wedList.Add(item);
                        if (MainWindow.main.wedApp_listbox.Items.Count < 2)
                            MainWindow.main.wedApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Thursday")
                    {
                        MainWindow.main.thuList.Add(item);
                        if (MainWindow.main.thuApp_listbox.Items.Count < 2)
                            MainWindow.main.thuApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Friday")
                    {
                        MainWindow.main.friList.Add(item);
                        if (MainWindow.main.friApp_listbox.Items.Count < 2)
                            MainWindow.main.friApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Saturday")
                    {
                        MainWindow.main.satList.Add(item);
                        if (MainWindow.main.satApp_listbox.Items.Count < 2)
                            MainWindow.main.satApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Sunday")
                    {
                        MainWindow.main.sunList.Add(item);
                        if (MainWindow.main.sunApp_listbox.Items.Count < 2)
                            MainWindow.main.sunApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                }
                else if (item.Start.Day == currDateTime.AddDays(3).Day)
                {
                    if (item.Start.DayOfWeek.ToString() == "Monday")
                    {
                        MainWindow.main.monList.Add(item);
                        if (MainWindow.main.monApp_listbox.Items.Count < 2)
                            MainWindow.main.monApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Tuesday")
                    {
                        MainWindow.main.tueList.Add(item);
                        if (MainWindow.main.tueApp_listbox.Items.Count < 2)
                            MainWindow.main.tueApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Wednesday")
                    {
                        MainWindow.main.wedList.Add(item);
                        if (MainWindow.main.wedApp_listbox.Items.Count < 2)
                            MainWindow.main.wedApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Thursday")
                    {
                        MainWindow.main.thuList.Add(item);
                        if (MainWindow.main.thuApp_listbox.Items.Count < 2)
                            MainWindow.main.thuApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Friday")
                    {
                        MainWindow.main.friList.Add(item);
                        if (MainWindow.main.friApp_listbox.Items.Count < 2)
                            MainWindow.main.friApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Saturday")
                    {
                        MainWindow.main.satList.Add(item);
                        if (MainWindow.main.satApp_listbox.Items.Count < 2)
                            MainWindow.main.satApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Sunday")
                    {
                        MainWindow.main.sunList.Add(item);
                        if (MainWindow.main.sunApp_listbox.Items.Count < 2)
                            MainWindow.main.sunApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                }
                else if (item.Start.Day == currDateTime.AddDays(4).Day)
                {
                    if (item.Start.DayOfWeek.ToString() == "Monday")
                    {
                        MainWindow.main.monList.Add(item);
                        if (MainWindow.main.monApp_listbox.Items.Count < 2)
                            MainWindow.main.monApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Tuesday")
                    {
                        MainWindow.main.tueList.Add(item);
                        if (MainWindow.main.tueApp_listbox.Items.Count < 2)
                            MainWindow.main.tueApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Wednesday")
                    {
                        MainWindow.main.wedList.Add(item);
                        if (MainWindow.main.wedApp_listbox.Items.Count < 2)
                            MainWindow.main.wedApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Thursday")
                    {
                        MainWindow.main.thuList.Add(item);
                        if (MainWindow.main.thuApp_listbox.Items.Count < 2)
                            MainWindow.main.thuApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Friday")
                    {
                        MainWindow.main.friList.Add(item);
                        if (MainWindow.main.friApp_listbox.Items.Count < 2)
                            MainWindow.main.friApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Saturday")
                    {
                        MainWindow.main.satList.Add(item);
                        if (MainWindow.main.satApp_listbox.Items.Count < 2)
                            MainWindow.main.satApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Sunday")
                    {
                        MainWindow.main.sunList.Add(item);
                        if (MainWindow.main.sunApp_listbox.Items.Count < 2)
                            MainWindow.main.sunApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                }
                else if (item.Start.Day == currDateTime.AddDays(5).Day)
                {
                    if (item.Start.DayOfWeek.ToString() == "Monday")
                    {
                        MainWindow.main.monList.Add(item);
                        if (MainWindow.main.monApp_listbox.Items.Count < 2)
                            MainWindow.main.monApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Tuesday")
                    {
                        MainWindow.main.tueList.Add(item);
                        if (MainWindow.main.tueApp_listbox.Items.Count < 2)
                            MainWindow.main.tueApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Wednesday")
                    {
                        MainWindow.main.wedList.Add(item);
                        if (MainWindow.main.wedApp_listbox.Items.Count < 2)
                            MainWindow.main.wedApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Thursday")
                    {
                        MainWindow.main.thuList.Add(item);
                        if (MainWindow.main.thuApp_listbox.Items.Count < 2)
                            MainWindow.main.thuApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Friday")
                    {
                        MainWindow.main.friList.Add(item);
                        if (MainWindow.main.friApp_listbox.Items.Count < 2)
                            MainWindow.main.friApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Saturday")
                    {
                        MainWindow.main.satList.Add(item);
                        if (MainWindow.main.satApp_listbox.Items.Count < 2)
                            MainWindow.main.satApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Sunday")
                    {
                        MainWindow.main.sunList.Add(item);
                        if (MainWindow.main.sunApp_listbox.Items.Count < 2)
                            MainWindow.main.sunApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                }
                else if (item.Start.Day == currDateTime.AddDays(6).Day)
                {
                    if (item.Start.DayOfWeek.ToString() == "Monday")
                    {
                        MainWindow.main.monList.Add(item);
                        if (MainWindow.main.monApp_listbox.Items.Count < 2)
                            MainWindow.main.monApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Tuesday")
                    {
                        MainWindow.main.tueList.Add(item);
                        if (MainWindow.main.tueApp_listbox.Items.Count < 2)
                            MainWindow.main.tueApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Wednesday")
                    {
                        MainWindow.main.wedList.Add(item);
                        if (MainWindow.main.wedApp_listbox.Items.Count < 2)
                            MainWindow.main.wedApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Thursday")
                    {
                        MainWindow.main.thuList.Add(item);
                        if (MainWindow.main.thuApp_listbox.Items.Count < 2)
                            MainWindow.main.thuApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Friday")
                    {
                        MainWindow.main.friList.Add(item);
                        if (MainWindow.main.friApp_listbox.Items.Count < 2)
                            MainWindow.main.friApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Saturday")
                    {
                        MainWindow.main.satList.Add(item);
                        if (MainWindow.main.satApp_listbox.Items.Count < 2)
                            MainWindow.main.satApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                    else if (item.Start.DayOfWeek.ToString() == "Sunday")
                    {
                        MainWindow.main.sunList.Add(item);
                        if (MainWindow.main.sunApp_listbox.Items.Count < 2)
                            MainWindow.main.sunApp_listbox.Items.Add(item.Subject + "  : " + item.Start.TimeOfDay.ToString());
                    }
                }
            }
            if (MainWindow.main.monList.Count >= 3) MainWindow.main.monApp_listbox.Items.Add("...more...");
            if (MainWindow.main.tueList.Count >= 3) MainWindow.main.tueApp_listbox.Items.Add("...more...");
            if (MainWindow.main.wedList.Count >= 3) MainWindow.main.wedApp_listbox.Items.Add("...more...");
            if (MainWindow.main.thuList.Count >= 3) MainWindow.main.thuApp_listbox.Items.Add("...more...");
            if (MainWindow.main.friList.Count >= 3) MainWindow.main.friApp_listbox.Items.Add("...more...");
            if (MainWindow.main.satList.Count >= 3) MainWindow.main.satApp_listbox.Items.Add("...more...");
            if (MainWindow.main.sunList.Count >= 3) MainWindow.main.sunApp_listbox.Items.Add("...more...");
        }

        //create an event/appointment
        private void createAppointment()
        {
            outlook.Application app = new outlook.Application();
            outlook.AppointmentItem appoinment = (outlook.AppointmentItem)app.CreateItem(outlook.OlItemType.olAppointmentItem);
            appoinment.Body = "Somethingg";
            appoinment.Importance = outlook.OlImportance.olImportanceNormal;
            appoinment.Save();   //  =  ((outlook._AppointmentItem)appoinment).Save();
        }

        private void deleteAppointment()
        {
            DateTime currDateTime = DateTime.Now.Date;

            outlook.AppointmentItem selectedItem = new outlook.AppointmentItem();
            MainWindow.main.monApp_listbox.SelectedItem.ToString();



            outlook._Application app = new outlook.Application();
            outlook._NameSpace ns = app.GetNamespace("MAPI");
            outlook.MAPIFolder mAPIFolder = ns.GetDefaultFolder(outlook.OlDefaultFolders.olFolderCalendar);
            
            foreach(outlook.AppointmentItem item in mAPIFolder.Items)
            {
                if (item.Equals(selectedItem))
                {
                    item.Delete();
                }
            }
        }
    }
}
