using GeneradorBingo;
using System;
using System.Collections.Generic;

class Program
{
    static Random random = new Random();
    static List<string[][]> cartonesGenerados = new List<string[][]>();
    static void Main()
    {
        int totalCartones = 300;
        int filas = 5;
        int columnas = 5;



        for (int i = 0; i < totalCartones; i++)
        {
            string[][] carton;
            bool esCartonUnico;
            bool filaRepetida = false;
            do
            {
                carton = GenerarCarton(filas, columnas);
                esCartonUnico = EsCartonUnico(carton, cartonesGenerados);
                filaRepetida=  FilasRepetidas(carton);

                if (filaRepetida)
                {
                    Console.WriteLine("Carton con fila repetida");
                    ImprimirCarton(carton, filas, columnas);
                }
                if (!esCartonUnico)
                {
                    Console.WriteLine("Carton repetido");
                    ImprimirCarton(carton, filas, columnas);
                }
            } while (!esCartonUnico || filaRepetida);

            cartonesGenerados.Add(carton);
            
            Exportar.CrearArchivo(carton, i + 1);

            Console.WriteLine($"Cartón {i + 1}:");
            ImprimirCarton(carton, filas, columnas);
            Console.WriteLine();
        }
    }


    static bool FilasRepetidas(string[][] carton)
    {
        if (cartonesGenerados.Count() == 0) return false;
        
        List<int> filasACuidar = new List<int>() { 0,4};

        foreach (var i in filasACuidar)
        {
            var matrizFila = carton[i].ToList();

            foreach(var item in cartonesGenerados)
            {
                var fila = item[i].ToList();

                var coincidencia = fila.Intersect(matrizFila).ToList();
                if (coincidencia.Count() > 4 ) 
                {
                    Console.WriteLine("Se repite");
                    
                    return true;
                } 
            }

        }
        return false;
    }

    static bool ColumnasRepetida(string[][] carton)
    {
        if (cartonesGenerados.Count() == 0) return false;

        for (int i = 0; i < 5; i++)
        {
            var matrizColumna = GetColumna(carton, i);

            var repite = cartonesGenerados
                .FirstOrDefault(n => GetColumna(n, i) == matrizColumna);

            if (repite != null) return true;
        }


        return false;
    }

    static string[] GetColumna(string[][] carton, int columna)
    {
        string[] matriz = new string[5];

        for (int fila = 0; fila < matriz.Length; fila++)
        {
            matriz[fila] = carton[fila][columna];
        }
        return matriz;
    }
    static string[][] GenerarCarton(int filas, int columnas)
    {
        string[][] carton = new string[filas][];

        for (int i = 0; i < filas; i++)
        {
            carton[i] = new string[columnas];

            for (int j = 0; j < columnas; j++)
            {
                if (i == 2 && j == 2)
                {
                    // Espacio libre en el centro
                    carton[i][j] = "libre";
                }
                else
                {
                    string letra = string.Empty;
                    int minNumero = j * 15 + 1;
                    int maxNumero = minNumero + 14;

                    int numero;
                    do
                    {
                        numero = random.Next(minNumero, maxNumero + 1);
                    } while (NumeroRepetido(carton, numero, j));

                    carton[i][j] = $"{letra}{numero:D2}";
                }
            }
        }

        return carton;
    }

    static bool NumeroRepetido(string[][] carton, int numero, int columna)
    {
        foreach (var fila in carton)
        {
            if (fila != null && fila[columna] != null && fila[columna].EndsWith(numero.ToString("D2")))
            {
                return true;
            }
        }
        return false;
    }

    static bool EsCartonUnico(string[][] carton, List<string[][]> cartonesGenerados)
    {
        foreach (var cartonGenerado in cartonesGenerados)
        {
            if (SonCartonesIguales(carton, cartonGenerado))
            {
                return false;
            }
        }
        return true;
    }

    static bool SonCartonesIguales(string[][] carton1, string[][] carton2)
    {
        var cartonLineal1 = new List<int>();
        var cartonLineal2 = new List<int>();
        foreach (var item in carton1)
        {
            cartonLineal1 = cartonLineal1.Concat(item.ToList().Select(x=> Convert.ToInt32(x=="libre" ? "0":x))).ToList();
        }

        foreach (var item in carton2)
        {
            cartonLineal2 = cartonLineal2.Concat(item.ToList().Select(x => Convert.ToInt32(x == "libre" ? "0":x))).ToList();
        }



        //Console.WriteLine("Coincidencia " + coincidencias);
       var repitencia = cartonLineal1.Intersect(cartonLineal2).Count();

        if(repitencia > 15)
        {
            Console.WriteLine($"Comparten {repitencia} letras");
        }

        return repitencia > 23;
    }

    static void ImprimirCarton(string[][] carton, int filas, int columnas)
    {
        for (int i = 0; i < filas; i++)
        {
            for (int j = 0; j < columnas; j++)
            {
                Console.Write(carton[i][j].PadLeft(6));
            }
            Console.WriteLine();
        }
    }
}
