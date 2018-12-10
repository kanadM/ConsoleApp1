using System;

namespace practice
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] arr1 = { "Banana", "Kiwi", "Mango", "Pinaple" }; //i
            string[] arr2 = { "Apple", "Banana", "Ornage", "Watermelon" }; //j
            int m = arr1.Length;
            int n = arr2.Length;

            findUncomman(arr1, arr2, m, n);
        }
        static void findUncomman(string[] arr1, string[] arr2, int m, int n)
        {
            int i = 0, j = 0;

            while (i < m && j < n)
            {
                if (arr1[i].CompareTo(arr2[j]) < 0)
                {
                    Console.WriteLine(arr1[i++]);
                    continue;
                }
                if (arr1[i].CompareTo(arr2[j]) > 0)
                {
                    Console.WriteLine(arr2[j++]);
                    continue;
                }
                else
                    j++; i++;
            }

            while (i < m)
            {
                Console.WriteLine(arr1[i++]);
            }

            while (j < n)
            {
                Console.WriteLine(arr2[j++]);
            }
        }

    }
}



