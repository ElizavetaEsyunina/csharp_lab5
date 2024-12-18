using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Laba5
{
    internal class Rooms
    {
        private int room_id;
        private int floor;
        private int number_ofBeds;
        private double accomodation_cost;
        private int category;

        public Rooms(int room_id, int floor, int number_ofBeds, double accomodation_cost, int category)
        {
            this.room_id = room_id;
            this.floor = floor;
            this.number_ofBeds = number_ofBeds;
            this.accomodation_cost = accomodation_cost;
            this.category = category;
        }

        public int RoomID
        {
            get { return room_id; }
            set { room_id = value; }
        }
        public int Floor
        {
            get { return floor; }
            set { floor = value; }
        }
        public int NumberOfBeds
        {
            get { return number_ofBeds; }
            set { number_ofBeds = value; }
        }
        public double AccomodationCost
        {
            get { return accomodation_cost; }
            set { accomodation_cost = value; }
        }
        public int Category
        {
            get { return category; }
            set { category = value; }
        }

        public override string ToString()
        {
            return $"{room_id, -10} | {floor, -5} | {number_ofBeds,-15} | {accomodation_cost, -20} | {category,-20}";
        }
    }
}
