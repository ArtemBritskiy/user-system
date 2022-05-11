using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1
{
    public class User : Person
    {
        public static void AddUser(string login, string password, string FirstName, string FullName, string email, string numberPhone)
        {
            StaticClass.persons[StaticClass.counterPersons] = new User();
            StaticClass.persons[StaticClass.counterPersons].UniqueNumber = StaticClass.counterPersons;
            StaticClass.persons[StaticClass.counterPersons].Login = login;
            StaticClass.persons[StaticClass.counterPersons].Password = password;
            StaticClass.persons[StaticClass.counterPersons].FirstName = FirstName;
            StaticClass.persons[StaticClass.counterPersons].FullName = FullName;
            StaticClass.persons[StaticClass.counterPersons].Email = email;
            StaticClass.persons[StaticClass.counterPersons].NumberPhone = numberPhone;
            StaticClass.persons[StaticClass.counterPersons].DataRegistration = DateTime.Now.ToString("dd.MM.yyyy");

            StaticClass.counterPersons++;

        }     
    }
}
