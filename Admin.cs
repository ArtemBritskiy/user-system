using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1
{
    public class Admin : User
    {
        public static void ChangeUserPanel3(int uniqueNumber, string firstName, string fullName, string email, string numberPhone, string password)
        {
            StaticClass.persons[uniqueNumber].FirstName = firstName;
            StaticClass.persons[uniqueNumber].FullName = fullName;
            StaticClass.persons[uniqueNumber].Email = email;
            StaticClass.persons[uniqueNumber].NumberPhone = numberPhone;
            StaticClass.persons[uniqueNumber].Password = password;
        }

        public static void ChangeUserPanel4(int uniqueNumber, string firstName, string fullName, string email, string numberPhone, string levelAccess)
        {
            StaticClass.persons[uniqueNumber].FirstName = firstName;
            StaticClass.persons[uniqueNumber].FullName = fullName;
            StaticClass.persons[uniqueNumber].Email = email;
            StaticClass.persons[uniqueNumber].NumberPhone = numberPhone;
            StaticClass.persons[uniqueNumber].LevelAccess = Convert.ToInt16(levelAccess);
        }
    }
}
