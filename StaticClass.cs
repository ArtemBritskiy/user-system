using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1
{
    internal class StaticClass
    {
        public static int quantityPersons = 100;                  // количество пользователей в системе
        public static int counterPersons = 0;                     // счётчик пользователей
        public static Person[] persons = new Person[quantityPersons];
      
    }
}
