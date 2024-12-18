using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Laba5
{
    internal class Booking
    {
        private int booking_id;
        private int client_id;
        private int room_id;
        private DateOnly booking_date;
        private DateOnly arrive_date;
        private DateOnly depart_date;

        public Booking (int booking_id, int client_id, int room_id, DateOnly booking_date, DateOnly arrive_date, DateOnly depart_date)
        {
            this.booking_id = booking_id;
            this.client_id = client_id;
            this.room_id = room_id;
            this.booking_date = booking_date;
            this.arrive_date = arrive_date;
            this.depart_date = depart_date;
        }

        public int BookingID
        {
            get { return booking_id; }
            set { booking_id = value; }
        }
        public int ClientID
        {
            get { return client_id; }
            set { client_id = value; }
        }
        public int RoomID
        {
            get { return room_id; }
            set { room_id = value; }
        }
        public DateOnly BookingDate
        {
            get { return booking_date; }
            set { booking_date = value; }
        }
        public DateOnly ArriveDate
        {
            get { return arrive_date; }
            set { arrive_date = value; }
        }
        public DateOnly DepartDate
        {
            get { return depart_date; }
            set { depart_date = value; }
        }

        public override string ToString()
        {
            return $"{booking_id, -15} | {client_id, -10} | {room_id, -10} | {booking_date, -20} | {arrive_date, -15} | {depart_date, -15}";
        }
    }
}
