using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    class Training
    {
        private DateTime kenshuuDate;
        private String courseNumber;
        private String courseName;
        private String department;
        private String staffCode;
        private String staffName;

        private Boolean lecturer;
        private Boolean text;
        private Boolean content;
        private Boolean continuation;
        private Boolean time;
        private Boolean day;
        private Boolean condition;

        private String lecturer_reason;
        private String text_reason;
        private String content_reason;
        private String continuation_reason;
        private String time_reason;
        private String day_reason;
        private String condition_reason;


        public string StaffName { get => staffName; set => staffName = value; }
        public string StaffCode { get => staffCode; set => staffCode = value; }
        public string Department { get => department; set => department = value; }
        public string CourseName { get => courseName; set => courseName = value; }
        public String CourseNumber { get => courseNumber; set => courseNumber = value; }
        public DateTime KenshuuDate { get => kenshuuDate; set => kenshuuDate = value; }
        public bool Lecturer { get => lecturer; set => lecturer = value; }
        public bool Text { get => text; set => text = value; }
        public bool Content { get => content; set => content = value; }
        public bool Continuation { get => continuation; set => continuation = value; }
        public bool Time { get => time; set => time = value; }
        public bool Day { get => day; set => day = value; }
        public bool Condition { get => condition; set => condition = value; }
        public string Lecturer_reason { get => lecturer_reason; set => lecturer_reason = value; }
        public string Text_reason { get => text_reason; set => text_reason = value; }
        public string Content_reason { get => content_reason; set => content_reason = value; }
        public string Continuation_reason { get => continuation_reason; set => continuation_reason = value; }
        public string Time_reason { get => time_reason; set => time_reason = value; }
        public string Day_reason { get => day_reason; set => day_reason = value; }
        public string Condition_reason { get => condition_reason; set => condition_reason = value; }
    }
}
