using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Excel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace Laba5
{
    internal class DataManager
    {
        private Application excelApp;
        private Workbook workbook;
        private Worksheet clientsSheet;
        private Worksheet bookingSheet;
        private Worksheet roomsSheet;

        public DataManager(string FilePath)
        {
            if (!File.Exists(FilePath)) throw new Exception($"Файла {FilePath} не существует");
            excelApp = new Application();
            excelApp.DisplayAlerts = false;
            workbook = excelApp.Workbooks.Open(FilePath);
            clientsSheet = (Worksheet) workbook.Sheets[1];
            bookingSheet = (Worksheet) workbook.Sheets[2];
            roomsSheet = (Worksheet) workbook.Sheets[3];
        }

        // получение данных из таблицы "Клиенты"
        public List<Clients> GetClients()
        {
            List<Clients> clients = new List<Clients>();
            for (int i = 2; i <= clientsSheet.UsedRange.Rows.Count; i++)
            {
                int client_id = (int)(clientsSheet.Cells[i, 1] as Range).Value2;
                string surname = (string)(clientsSheet.Cells[i, 2] as Range).Value2;
                string name = (string)(clientsSheet.Cells[i, 3] as Range).Value2;
                string patronymic = (string)(clientsSheet.Cells[i, 4] as Range).Value2;
                string address = (string)(clientsSheet.Cells[i, 5] as Range).Value2;
                Clients client = new Clients(client_id, surname, name, patronymic, address);
                clients.Add(client);
            }
            return clients;
        }
        // получение данных из таблицы "Бронирование"
        public List<Booking> GetBookings()
        {
            List<Booking> bookings = new List<Booking>();
            for (int i = 2; i <= bookingSheet.UsedRange.Rows.Count; i++)
            {
                int booking_id = (int)(bookingSheet.Cells[i, 1] as Range).Value2;
                int client_id = (int)(bookingSheet.Cells[i, 2] as Range).Value2;
                int room_id = (int)(bookingSheet.Cells[i, 3] as Range).Value2;
                DateOnly booking_date = DateOnly.FromDateTime(DateTime.FromOADate((double)(bookingSheet.Cells[i, 4] as Range).Value2));
                DateOnly arrive_date = DateOnly.FromDateTime(DateTime.FromOADate((double)(bookingSheet.Cells[i, 5] as Range).Value2));
                DateOnly depart_date = DateOnly.FromDateTime(DateTime.FromOADate((double)(bookingSheet.Cells[i, 6] as Range).Value2));
                Booking booking = new Booking(booking_id, client_id, room_id, booking_date, arrive_date, depart_date);
                bookings.Add(booking);
            }
            return bookings;
        }
        // получение данных из таблицы "Номера"
        public List<Rooms> GetRooms()
        {
            List<Rooms> rooms = new List<Rooms>();
            for (int i = 2; i <= roomsSheet.UsedRange.Rows.Count; i++)
            {
                int room_id = (int)(roomsSheet.Cells[i, 1] as Range).Value2;
                int floor = (int)(roomsSheet.Cells[i, 2] as Range).Value2;
                int number_ofBeds = (int)(roomsSheet.Cells[i, 3] as Range).Value2;
                double cost = (double)(roomsSheet.Cells[i, 4] as Range).Value2;
                int category = (int)(roomsSheet.Cells[i, 5] as Range).Value2;
                Rooms room = new Rooms(room_id, floor, number_ofBeds, cost, category);
                rooms.Add(room);
            }
            return rooms;
        }

        //просмотр базы данных
        public void ViewClients()
        {
            Console.WriteLine($"{"Код клиента",-10}| {"Фамилия",-15} | {"Имя",-15} | {"Отчество",-15} | {"Место жительства",-10}");
            List<Clients> clients = GetClients();
            foreach (Clients client in clients)
            {
                Console.WriteLine(client);
            }
        }
        public void ViewBooking()
        {
            Console.WriteLine($"{"Код бронирования",-15}| {"Код клиента",-10}| {"Код номера",-10} | {"Дата бронирования",-20} | {"Дата заезда",-15} | {"Дата выезда",-15}");
            List<Booking> bookings = GetBookings();
            foreach (Booking booking in bookings)
            {
                Console.WriteLine(booking);
            }
        }
        public void ViewRooms()
        {
            Console.WriteLine($"{"Код номера",-10} | {"Этаж",-5} | {"Число мест",-15} | {"Стоимость проживания",-20} | {"Категория",-20}");
            List<Rooms> rooms = GetRooms();
            foreach (Rooms room in rooms)
            {
                Console.WriteLine(room);
            }
        }

        // 1) Определите общую стоимость проживания за сутки в номерах категории 5, забронированных клиентами из г.Уфа с 1 по 16 июня включительно.
        public void TotalCost()
        {
            List<Clients> clients = GetClients();
            List<Booking> bookings = GetBookings();
            List<Rooms> rooms = GetRooms();
            DateOnly start = new DateOnly(2019, 6, 1);
            DateOnly end = new DateOnly(2019, 6, 16);
            var totalCost = (from booking in bookings
                             join client in clients on booking.ClientID equals client.Client_ID
                             join room in rooms on booking.RoomID equals room.RoomID
                             where room.Category == 5 && client.Address == "г. Уфа" && booking.BookingDate >= start &&  booking.BookingDate <= end
                             select room.AccomodationCost).Sum();
            Console.WriteLine("Общая стоимость проживания: " + totalCost);
        }
        // 2) вывести ФИО всех клиентов, проживающих в Уфе
        public void ClientsFromUfa()
        {
            List<Clients> clients = GetClients();
            var ufa_clients = (from client in clients
                       where client.Address == "г. Уфа"
                       select new { fio = $"{client.Surname} {client.Name} {client.Patronymic}" });
            Console.WriteLine("ФИО всех клиентов из г. Уфа:");
            foreach (var client in ufa_clients)
            {
                Console.WriteLine(client.fio);
            }
            Console.WriteLine("\nОбщее число клиентов из Уфы: " + ufa_clients.Count());
        }
        // 3) Вывести ФИО всех клиентов из Уфы, прибывших с 11.07.2019 по 13.07.2019 включительно. Отсортировать в порядке возрастания кода бронирования
        public void ArriveOnThisDate()
        {
            List<Clients> clients = GetClients();
            List<Booking> bookings = GetBookings();
            DateOnly arrive_date1 = new DateOnly(2019, 07, 11);
            DateOnly arrive_date2 = new DateOnly(2019, 07, 13);
            var arrive_onThisDate = (from booking in bookings
                                     join client in clients on booking.ClientID equals client.Client_ID
                                     where booking.ArriveDate >= arrive_date1 && booking.ArriveDate <= arrive_date2 && client.Address == "г. Уфа"
                                     orderby booking.BookingID ascending
                                     select new { fio = $"{client.Surname} {client.Name} {client.Patronymic}" });
            Console.WriteLine("ФИО всех клиентов из Уфы, прибывших с 11.07.2019 по 13.07.2019 включительно:");
            foreach (var client in arrive_onThisDate)
            {
                Console.WriteLine(client.fio);
            }
        }
        // 4) Определить макс. стоимость проживания за сутки в номерах категории 1, забронированных клиентами из Уфы с 1 по 16 июня включительно
        public void MaxCost()
        {
            List<Clients> clients = GetClients();
            List<Booking> bookings = GetBookings();
            List<Rooms> rooms = GetRooms();
            DateOnly start = new DateOnly(2019, 6, 1);
            DateOnly end = new DateOnly(2019, 6, 16);
            var max_cost = (from room in rooms
                            join booking in bookings on room.RoomID equals booking.RoomID
                            join client in clients on booking.ClientID equals client.Client_ID
                            where room.Category == 1 && client.Address == "г. Уфа" && booking.BookingDate >= start && booking.BookingDate <= end
                            select room.AccomodationCost).Max();
            Console.WriteLine("Макс. стоимость проживания: " + max_cost);
        }

        public void Close()
        {
            // Закрываем рабочую книгу
            workbook.Close();

            // Освобождаем объекты
            Marshal.ReleaseComObject(clientsSheet);
            Marshal.ReleaseComObject(bookingSheet);
            Marshal.ReleaseComObject(roomsSheet);
            Marshal.ReleaseComObject(workbook);

            // Закрываем Excel
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);

            // Вызываем сборщик мусора
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
