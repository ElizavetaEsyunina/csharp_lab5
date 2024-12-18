using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Laba5
{
    internal class Clients
    {
        private int client_id;
        private string surname;
        private string name;
        private string patronymic;
        private string address;

        public Clients(int client_id, string surname, string name, string patronymic, string address)
        {
            this.client_id = client_id;
            this.surname = surname;
            this.name = name;
            this.patronymic = patronymic;
            this.address = address;
        }

        public int Client_ID
        {
            get { return client_id; }
            set { client_id = value; }
        }
        public string Surname
        {
            get { return surname; }
            set { surname = value; }
        }
        public string Name
        {
            get { return name; }
            set { name = value; }
        }
        public string Patronymic
        {
            get { return patronymic; }
            set { patronymic = value; }
        }
        public string Address
        {
            get { return address; }
            set { address = value; }
        }

        public override string ToString()
        {
            return $"{client_id, -10} | {surname, -15} | {name, -15} | {patronymic, -15} | {address, -10}";
        }
    }
}
