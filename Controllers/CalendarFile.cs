using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

namespace Sitecore.SharedSource.AddToCalendar
{

    public class CalendarFile
    {
        #region Constructors
        public CalendarFile(string p_title, DateTime p_startDate, DateTime p_endDate)
        {
            if (p_startDate > p_endDate) throw new Exception("VCSFile: Attept to initialize a calendar event with start date after end date");
            _summary = p_title;
            _startDate = p_startDate;
            _endDate = p_endDate;
        }
        #endregion
        #region Enums
        public enum CFStatus { Free, Busy, Tentative, OutOfTheOffice };
        #endregion
        #region Public Methods
        override public string ToString()
        {
            StringBuilder calEvent = new StringBuilder();
            calEvent.AppendLine("BEGIN:VCALENDAR");
            calEvent.AppendLine("VERSION:2.0");
            calEvent.AppendLine("PRODID:-//hacksw/handcal//NONSGML v1.0//EN");
            calEvent.AppendLine("BEGIN:VEVENT");
            switch (status)
            {
                case CFStatus.Busy:
                    calEvent.AppendLine("X-MICROSOFT-CDO-BUSYSTATUS:BUSY");
                    break;
                case CFStatus.Free:
                    calEvent.AppendLine("TRANSP:TRANSPARENT");
                    break;
                case CFStatus.Tentative:
                    calEvent.AppendLine("STATUS:TENTATIVE");
                    break;
                case CFStatus.OutOfTheOffice:
                    calEvent.AppendLine("X-MICROSOFT-CDO-BUSYSTATUS:OOF");
                    break;
                default:
                    throw new Exception("Invalid CFStatus");
            }
            if (allDayEvent)
            {
                calEvent.AppendLine("DTSTART;VALUE=DATE:" + startDate.ToCFDateOnlyString());
                calEvent.AppendLine("DTEND;;VALUE=DATE:" + endDate.ToCFDateOnlyString());
            }
            else
            {
                calEvent.AppendLine("DTSTART:" + startDate.ToCFString());
                calEvent.AppendLine("DTEND:" + endDate.ToCFString());
            }
            calEvent.AppendLine("SUMMARY:" + summary);
            if (location != "") calEvent.AppendLine("LOCATION:" + location);
            if (description != "") calEvent.AppendLine("DESCRIPTION:" + description);
            calEvent.AppendLine("END:VEVENT");
            calEvent.AppendLine("END:VCALENDAR");
            return calEvent.ToString();
        }
        #endregion
        #region Accessors and Mutators
        public string summary
        {
            get { return _summary; }
            set { _summary = value; }
        }
        public DateTime startDate
        {
            get { return _startDate; }
            //set { _startDate = value; }
        }
        public DateTime endDate
        {
            get { return _endDate; }
            //set { _endDate = value; }
        }
        public string location
        {
            get { return _location; }
            set { _location = value; }
        }
        public string description
        {
            get { return _description; }
            set { _description = value; }
        }
        public bool allDayEvent
        {
            get { return _bookFullDays; }
            set { _bookFullDays = value; }
        }
        public CFStatus status
        {
            get { return _status; }
            set { _status = value; }
        }
        #endregion
        #region Internal Fields
        //mandatory fields initialized in the constructor
        private string _summary;
        private DateTime _startDate;
        private DateTime _endDate;
        //fields with default values
        //private ShowStatusAs _showStatusAs = ShowStatusAs.Busy;
        private CFStatus _status = CFStatus.Busy;
        private string _location = "";
        private string _description = "";
        private bool _bookFullDays = false;
        #endregion
    }


    public static class CalendarFileExtentions
    {
        public static string ToCFString(this DateTime val)
        {
            //format: YYYYMMDDThhmmssZ where YYYY = year, MM = month, DD = date, T = start of time character, hh = hours, mm = minutes,
            //ss = seconds, Z = end of tag character. The entire tag uses Greenwich Mean Time (GMT)
            //return val.ToUniversalTime().ToString("yyyyMMddThhmmssZ");
            return val.ToUniversalTime().ToString("yyyyMMdd\\THHmmss\\Z");
        }
        public static string ToCFDateOnlyString(this DateTime val)
        {
            //format: YYYYMMDD where YYYY = year, MM = month, DD = date, T = start of time character, hh = hours, mm = minutes,
            //ss = seconds, Z = end of tag character. The entire tag uses Greenwich Mean Time (GMT)
            return val.ToUniversalTime().ToString("yyyyMMdd");
        }
        public static string ToCFString(this string str)
        {
            return str.Replace("r", "=0D").Replace("n", "=0A");
        }
        public static string ToOneLineString(this string str)
        {
            return str.Replace("r", " ").Replace("n", " ");
        }
    }
}