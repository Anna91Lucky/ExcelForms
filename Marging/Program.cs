using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Marging
{
    internal class Program
    {
        static void Main(string[] args)
        {

            string[] soft = { "Microsoft", "Google", "Apple" };
            string[] hard = { "Apple", "IBM", "Samsung" };

            // объединение последовательностей
            var result = soft.Union(hard);

            foreach (string s in result)
                Console.WriteLine(s);
            
            Console.ReadKey();
        }
    }
}
