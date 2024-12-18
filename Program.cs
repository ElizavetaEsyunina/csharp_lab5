using Laba5;
internal class Program
{
    private static void Main(string[] args)
    {
        try
        {
            string filepath = "D:\\2 курс\\Языки программирования\\Лабы\\Лаба 5\\LR5-var9.xls";
            Console.WriteLine("Вести протоколирование действий в новом файле или дописывать в уже существующий?\n" +
                "Введите 'новый' или 'дописать'");
            string choice = Console.ReadLine().ToLower();
            Logging logger = new Logging();
            string log_file;
            try
            {
                if (choice == "новый")
                {
                    Console.WriteLine("Введите путь к новому файлу для протоколирования:");
                    log_file = Console.ReadLine();

                    // Удаляем существующий файл, если он есть
                    if (File.Exists(log_file))
                    {
                        Console.WriteLine($"Файл {log_file} уже существует. Он будет удален и вместо него будет записан новый файл");
                        File.Delete(log_file);
                    }
                    logger.Log(log_file, "Начало нового сеанса.");
                    Console.WriteLine("Нажмите любую клавишу для продолжения...");
                    Console.ReadKey();
                }
                else if (choice == "дописать")
                {
                    Console.WriteLine("Введите путь к файлу для дописывания:");
                    log_file = Console.ReadLine();

                    // Проверяем, существует ли файл
                    if (!File.Exists(log_file))
                    {
                        Console.WriteLine($"Файла {log_file} не существует. Будет создан новый файл");
                        logger.Log(log_file, "Начало нового сеанса.");
                    }
                    else
                    {
                        Console.WriteLine($"Файл {log_file} существует. Данные будут дописываться в него");
                        logger.Log(log_file, "Начало сеанса (дописываем в существующий файл).");
                    }
                    Console.WriteLine("Нажмите любую клавишу для продолжения");
                    Console.ReadKey();
                }
                else
                {
                    Console.WriteLine("Некорректный выбор. Завершение работы программы.");
                    return;
                }

                while (true)
                {
                    Console.Clear();
                    Console.WriteLine("Выберите действие:\n" +
                        "1) Просмотр базы данных (таблицы \"клиенты\", \"бронирование\", \"номера\")\n" +
                        "\n<<<Реализация 4-х запросов>>>\n\n" +
                        "2) Определение общей стоимости проживания за сутки в номерах категории 5, забронированных клиентами из г.Уфа с 1 по 16 июня включительно\n" +
                        "3) Вывод ФИО всех клиентов, проживающих в Уфе\n" +
                        "4) Вывод ФИО всех клиентов из Уфы, прибывших с 11.07.2019 по 13.07.2019 включительно, в порядке возрастания кода бронирования\n" +
                        "5) Определение макс. стоимости проживания за сутки в номерах категории 1, забронированных клиентами из Уфы с 1 по 16 июня включительно\n" +
                        "\n0 - выход");
                    short choise = short.Parse(Console.ReadLine());
                    switch (choise)
                    {
                        case 1:
                            Console.Clear();
                            logger.Log(log_file, "Просмотр базы данных");
                            Console.WriteLine("Просмотр базы данных");
                            DataManager data = new DataManager(filepath);
                            Console.WriteLine("\nКлиенты:");
                            data.ViewClients();
                            Console.WriteLine("\nБронирование:");
                            data.ViewBooking();
                            Console.WriteLine("\nНомера:");
                            data.ViewRooms();
                            data.Close();
                            logger.Log(log_file, "Завершение действия");
                            break;

                        case 2:
                            Console.Clear();
                            logger.Log(log_file, "Выполнение запроса 1");
                            Console.WriteLine("Определение общей стоимости проживания за сутки в номерах категории 5, забронированных клиентами из г.Уфа с 1 по 16 июня включительно");
                            DataManager data6 = new DataManager(filepath);
                            data6.TotalCost();
                            data6.Close();
                            logger.Log(log_file, "Завершение действия");
                            break;
                        case 3:
                            Console.Clear();
                            logger.Log(log_file, "Выполнение запроса 2");
                            Console.WriteLine("Вывод ФИО всех клиентов, проживающих в Уфе");
                            DataManager data7 = new DataManager(filepath);
                            data7.ClientsFromUfa();
                            data7.Close();
                            logger.Log(log_file, "Завершение действия");
                            break;
                        case 4:
                            Console.Clear();
                            logger.Log(log_file, "Выполнение запроса 3");
                            Console.WriteLine("Вывод ФИО всех клиентов из Уфы, прибывших с 11.07.2019 по 13.07.2019 включительно, в порядке возрастания кода бронирования");
                            DataManager data8 = new DataManager(filepath);
                            data8.ArriveOnThisDate();
                            data8.Close();
                            logger.Log(log_file, "Завершение действия");
                            break;
                        case 5:
                            Console.Clear();
                            logger.Log(log_file, "Выполение запроса 4");
                            Console.WriteLine("Определение макс. стоимости проживания за сутки в номерах категории 1, забронированных клиентами из Уфы с 1 по 16 июня включительно");
                            DataManager data9 = new DataManager(filepath);
                            data9.MaxCost();
                            data9.Close();
                            logger.Log(log_file, "Завершение действия");
                            break;

                        case 0:
                            Console.Clear();
                            Console.WriteLine("Работа программы завершена");
                            logger.Log(log_file, "Завершение работы программы");
                            return;

                        default:
                            Console.Clear();
                            Console.WriteLine("Неверный выбор. Попробуйте ещё раз");
                            break;
                    }
                    Console.WriteLine("Нажмите любую клавишу для продолжения...");
                    Console.ReadKey();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
            return;
        }
    }
}