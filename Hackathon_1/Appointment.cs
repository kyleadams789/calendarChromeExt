using System;
using System.Collections.Generic;

public class Appointment
{
    public Appointment() { }


        public string Date;
        public string apptName;
        public List<string> emailList = new List<string>();

    public void setDate(string newDate)
    {
        this.Date = newDate;
    }
    public void setAppt(string appt)
    {
        this.apptName = appt;
    }
    public void fillEmails(string newEmail)
    {
        emailList.Add(newEmail);
    }

}