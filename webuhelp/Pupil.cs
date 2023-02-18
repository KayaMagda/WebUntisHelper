namespace webuhelp
{
    public class Pupil
    {
        public int ID { get; set; }
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string Class { get; set; }

        public string GetFullNameGermanBurocratic()
        {
            return LastName + " " + FirstName;
        }
    }    
}
