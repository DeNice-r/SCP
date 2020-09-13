namespace Schedule_Calculator_Pro
{
    public class Subject
    {
        public string subjectName { get; set; }
        public string relatedAud { get; set; } = "";
        public Subject(string subjectName)
        {
            this.subjectName = subjectName;
        }
    }
}
