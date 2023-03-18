namespace webuhelp
{
    public class PupilData
    {
        public bool IsExcused { get; set; }
        public string Date { get; set; }
        public string Weekday { get; set; }
        public int LessonNr { get; set; }
        public string Teacher { get; set; }
        public string Lesson { get; set; }
        public int MissingHour { get; set; }
        public int MissingMinute { get; set; }
        public string Reason { get; set; }
        public string? MissingText { get; set; }
        public string? Text { get; set; }
    }
}
