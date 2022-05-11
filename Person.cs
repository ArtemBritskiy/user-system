using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1
{
    public class Person
    {
        int uniqueNumber;       // уникальный номер
        string firstName;       // имя пользователя
        string fullName;        // фамилия пользователя
        string login;
        string password;        // пароль
        string email;           // e-mail
        string numberPhone;     // номер телефона
        string dataRegistration;
        int levelAccess;        // уровень допуступа
        public Person()
        {

        }

        public int UniqueNumber
        {
            get { return uniqueNumber; }   // возвращаем значение свойства

            set { uniqueNumber = value; }
        }

        public int LevelAccess
        {
            get { return levelAccess; }   // возвращаем значение свойства

            set { levelAccess = value; }
        }

        public string FirstName
        {
            get { return firstName; }   // возвращаем значение свойства

            set { firstName = value; }   // устанавливаем новое значение свойства
        }

        public string FullName
        {
            get { return fullName; }   // возвращаем значение свойства

            set { fullName = value; }   // устанавливаем новое значение свойства
        }

        public string Email
        {
            get { return email; }   // возвращаем значение свойства

            set { email = value; }   // устанавливаем новое значение свойства
        }

        public string Password
        {
            get { return password; }   // возвращаем значение свойства

            set { password = value; }   // устанавливаем новое значение свойства
        }

        public string Login
        {
            get { return login; }   // возвращаем значение свойства

            set { login = value; }   // устанавливаем новое значение свойства
        }

        public string NumberPhone
        {
            get { return numberPhone; }   // возвращаем значение свойства

            set { numberPhone = value; }   // устанавливаем новое значение свойства
        }


        public string DataRegistration
        {
            get { return dataRegistration; }   // возвращаем значение свойства

            set { dataRegistration = value; }   // устанавливаем новое значение свойства
        }

        public static int FindLevelAccess(string login)
        {
            int levelAccess = -1;
            for (int i = 0; i < StaticClass.counterPersons; i++)
            {
                if (StaticClass.persons[i].Login == login)
                {
                    levelAccess = StaticClass.persons[i].LevelAccess;
                }
            }
            return levelAccess;
        }

        public static int FindUniqueNumber(string login)
        {
            int uniqueNumber = -1;
            for (int i = 0; i < StaticClass.counterPersons; i++)
            {
                if (StaticClass.persons[i].Login == login)
                {
                    uniqueNumber = StaticClass.persons[i].UniqueNumber;
                }
            }
            return uniqueNumber;
        }
    }
}
