using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Laba5
{
    // для протоколирования действий
    internal class Logging
    {
        public void Log(string filepath, string message)
        {
            if (!filepath.Contains(".txt")) throw new Exception("Неверный формат файла для протоколирования");
            using (StreamWriter writer = new StreamWriter(File.Open(filepath, FileMode.Append)))
            {
                writer.WriteLine($"{DateTime.Now}: {message}");
            }
        }
    }
}
