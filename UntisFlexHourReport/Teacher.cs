namespace UntisFlexHourReport
{
    public class Teacher
    {
        public Teacher(string shortName, string firstName, string lastName, decimal actualHours)
        {
            NameCode = shortName;
            FirstName = firstName;
            LastName = lastName;
            ActualHours = actualHours;
        }

        public string NameCode { get; }
        public string FirstName { get; }
        public string LastName { get; }
        public decimal ActualHours { get; }
    }
}
