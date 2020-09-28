using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelProjectMvc.Models
{
    public class ListPerson
    {
        private List<PersonModel> _persons = new List<PersonModel>();
        public List<PersonModel> Persons
        {
            get { return _persons; }
        }
        public void AddPerson(PersonModel model)
        {
            _persons.Add(new PersonModel() { Name = model.Name, Surname = model.Surname, Email = model.Email, Address = model.Address });

        }
    }
}