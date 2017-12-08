using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RendToOutlook
{
    public class Room
    {
        private string number;
        private string name;
        private string email;
        private string rendezvousName;

        public string Number
        {
            get { return number; }
            set { number = value; }
        }

        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        public string Email
        {
            get { return email; }
            set { email = value; }
        }

        public string RendezvousName
        {
            get { return rendezvousName; }
            set { rendezvousName = value; }
        }

        public Room()
        {
            Number = null;
            Name = null;
            Email = null;
            RendezvousName = null;
        }
    }
}