using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookTest
{
   public class EventBody
    {
        public string ContentType { get; set; }
        public string Content { get; set; }
      
    }
    public class EventModel
    {
        public string Id { get; set; }
        public string Subject { get; set; }
        public DateTimeTimeZones Start { get; set; }
        public DateTimeTimeZones End { get; set; }
        public string ShowAs { get; set; }
        public bool IsAllDay { get; set; }
        public EventBody Body { get; set; }
        public bool IsReminderOn { get; set; }
    }

    public class DateTimeTimeZones
    {
        public DateTime DateTime { get; set; }
        public string TimeZone { get; set; }
    }
    public class TaskTarea
    {
  
        public int Id { get; set; }
        public string UserId { get; set; }
        public string Title { get; set; }
        public int? CtaId { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public Guid? CustomerId { get; set; }
        public string Notes { get; set; }
        public int Status { get; set; }
        public int RescheduleId { get; set; }
        public string RescheduleReason { get; set; }
      public bool IsEditable { get; set; }
        public string ExternalId { get; set; }
        public int Type { get; set; }
    }
}
