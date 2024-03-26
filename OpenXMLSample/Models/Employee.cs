namespace OpenXMLSample.Models
{
    public class EmployeeIndex
    {
        public const int userId = 2;
        public const int name = 4;
        public const int objective = 6;
        public const int educations = 8;
        public const int certifications = 10;
        public const int hobbies = 12;
        public const int skills = 14;
        public const int avatarFileId = 16;
        public const int languageId = 18;
        public const int isActive = 20;
        public const int hightLight = 22;
        public const int strength = 24;
        public const int urlGithub = 26;
        public const int presenter = 28;
        public const int honorsAndAwards = 30;
        public const int position = 32;
    }
    public static class EmployeeHelper
    {
        public static string GetEmployeeProp(this string input) => input.Split(":")[1].Trim();
    }
    public class Employee
    {
        public string userId { get; set; }
        public string name { get; set; }
        public string objective { get; set; }
        public string educations { get; set; }
        public string certifications { get; set; }
        public string hobbies { get; set; }
        public string skills { get; set; }
        public long avatarFileId { get; set; }
        public long? languageId { get; set; }
        public bool? isActive { get; set; }
        public string hightLight { get; set; }
        public string strength { get; set; }
        public string urlGithub { get; set; }
        public string presenter { get; set; }
        public string honorsAndAwards { get; set; }
        public string position { get; set; }
    }
}
