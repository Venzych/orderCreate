using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace orderCreateFinal.Models
{
    public class User
    {
        public long Id { get; set; }
        public string Name { get; set; }

        public User(long id, string name)
        {
            Id = id;
            Name = name;
        }
    }
}
