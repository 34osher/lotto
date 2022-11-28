using Bytescout.Spreadsheet;
using System;
using System.Diagnostics.Metrics;
using System.Reflection.Metadata;

namespace ReadFromExcel // Note: actual namespace depends on the project name.
{
    internal class Program
    {
       // static int Simmolarity;
        static int[] RandomNun(int[] p)
        {
            Random rnd = new Random();
            for (int i = 0; i < 7 - 1; i++)
            {
                p[i] = rnd.Next(1, 38);

            }
            p[6] = rnd.Next(1, 8);
            return p;
        }
        static int Check1(int[] w) //onle one number in the array
        {
            
            for (int i = 0; i < w.Length-1; i++)
            {
                for (int j = 1; (j + i) <= (w.Length - 1); j++)
                {
                    if (w[i] == w[i+ j])
                    {
                        return 0;
                    }

                }
                
            }
            return 1;

        }
       
        static int Check2(int[] r)
        {
            //Simmolarity = 0;
            int Count = 0;
            int te = 0;
            Spreadsheet doc = new Spreadsheet();
            doc.LoadFromFile(@"C:\Users\34osh\OneDrive\שולחן העבודה\Lotto.xlsx");
            Worksheet ws = doc.Workbook.Worksheets.ByName("Lotto");
            
            for (int i = 0; i < 4103; i++)
            {
                Count = 0;
                for (int j = 0; j < 7; j++)
                {
                    te = Check3(ws.Cell(i, j).ValueAsInteger, r);
                    Count = Count + te;

                    if (Count >=5)
                    {
                        return 0;

                    }

                }
            }
            doc.Close();
            Console.ReadKey();

            return 1;
        }

        static int Check3(int v,int[] l)
        {
           

            for (int z = 0; z < 7; z++)
            {

                if (v == l[z])
                {
                    return 1;
                    
                }
            }
            return 0;
        }
        static void Main()
        {
            int i, ch1 = 0,ch2 = 0,m=0;
            int[] a = new int[7];

            while (ch1 == 0 || ch2 == 0 )
            {
                m++;
                Console.WriteLine(" m = "+m);
                
                a = RandomNun(a);
                
                ch1 = Check1(a);              //onle one number in the array
                if (ch1 == 0) continue;
                Console.WriteLine("ch1 ="+ch1);

                Console.WriteLine("the ran after ch1");
                for (i = 0; i < 7; i++)
                {
                    Console.Write(" " + a[i]);

                }
                Console.WriteLine(" ");

                ch2 = Check2(a);
                Console.WriteLine("ch2 =" + ch2);
            }





            for (i = 0; i < a.Length; i++)
            {
                Console.Write(" "+a[i]);

            }
            Console.WriteLine("" );
            Console.Write("its aper " + ch2);


        }
    }
}