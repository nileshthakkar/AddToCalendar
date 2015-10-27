using Glass.Mapper.Sc;
using Sitecore.Data.Items;
using Sitecore.Mvc.Presentation;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Sitecore.SharedSource.AddToCalendar
{
    public class AddToCalendarController : Controller
    {        
        // GET: AddToCalendar
        [HttpGet]
        public ActionResult Index()
        {            
            Session["ID"] = Datasource.ID;   
            return View(Datasource);
        }      

        public void Download()
        {
            Item calenderItem = Sitecore.Context.Database.GetItem(Session["ID"].ToString());
            DateTime chatStartDate = DateTime.Today;
            DateTime chatEndDate = DateTime.Today;
            string strTimings = string.Empty;
            NameValueCollection collection = Sitecore.Web.WebUtil.ParseUrlParameters(calenderItem["Event Dates"]);
            if (collection.Count > 0)
            {
                for (int i = 0; i <= 6; i++)
                {
                    foreach (string day in collection.AllKeys)
                    {
                        if (chatStartDate.Date.AddDays(i).DayOfWeek.ToString() == day)
                        {
                            chatStartDate = chatStartDate.Date.AddDays(i);
                            strTimings = collection[day];
                            i = 7;
                            break;
                        }
                    }
                }
            }

            string[] chatTimings = strTimings.Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries);
            if (chatTimings.Length > 0)
            {
                if (chatTimings[0] != null)
                {
                    TimeSpan ts = new TimeSpan(int.Parse(chatTimings[0]), 0, 0);
                    chatStartDate = chatStartDate.Date + ts;
                }

                if (chatTimings[1] != null)
                {
                    TimeSpan ts2 = new TimeSpan(int.Parse(chatTimings[1]), 0, 0);
                    chatEndDate = chatStartDate.Date + ts2;
                }

                if (calenderItem["Time Zone"] != null)
                {
                    Item timeZoneItem = Sitecore.Context.Database.GetItem(calenderItem["Time Zone"]);

                    chatStartDate = convertDateToTimeZone(chatStartDate, timeZoneItem["Name"], "UTC");
                    chatEndDate = convertDateToTimeZone(chatEndDate, timeZoneItem["Name"], "UTC");

                    CalendarFile calendarFile = new CalendarFile(calenderItem["Subject"], chatStartDate, chatEndDate);
                    calendarFile.location = calenderItem["Location"];
                    calendarFile.description = calenderItem["Body"];

                    this.Response.Clear();
                    Response.Expires = 0;
                    this.Response.Buffer = true;
                    Response.ContentType = "text/calendar";
                    Response.AddHeader("Content-Disposition", "inline; filename=" + calenderItem.Name + ".ics");
                    this.Response.Write(calendarFile.ToString());
                    this.Response.Flush();
                    this.Response.End();
                }
            }
        }

        public DateTime convertDateToTimeZone(DateTime nextDate, string fromTimeZone, string toTimeZone)
        {
            TimeZoneInfo cstZone = TimeZoneInfo.FindSystemTimeZoneById(fromTimeZone);
            TimeZoneInfo cstZone2 = TimeZoneInfo.FindSystemTimeZoneById(toTimeZone);
            nextDate = DateTime.SpecifyKind(nextDate, DateTimeKind.Unspecified);
            DateTime cstDateTime = TimeZoneInfo.ConvertTime(nextDate, cstZone, cstZone2);
            return cstDateTime;
        }
        /// <summary>  
        /// Get either the data source or the context item  
        /// </summary>  
        public Item Datasource
        {
            get
            {
                var dataSource = Sitecore.Context.Item;
                string datasourceId = RenderingContext.Current.Rendering.DataSource;
                if (Sitecore.Data.ID.IsID(datasourceId))
                {
                    dataSource = Sitecore.Context.Database.GetItem(datasourceId);
                }

                return dataSource;
            }
        }  
    }
}