using System.Collections.Generic;

namespace webuhelp
{
    public class Pupil
    {
        public int ID { get; set; }
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string Class { get; set; }
        public List<PupilData> Data { get; set; }
        public int A { get; set; }
        public int N { get; set; }
        public int B { get; set; }
        public int V { get; set; }

        public string GetFullNameGermanBurocratic()
        {
            return LastName + " " + FirstName;
        }

        public string GetFileName()
        {
            return LastName + "_" + FirstName;
        }
    }    
}
