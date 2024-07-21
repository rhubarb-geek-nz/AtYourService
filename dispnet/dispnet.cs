/***********************************
 * Copyright (c) 2024 Roger Brown.
 * Licensed under the MIT License.
 ****/

using System;

namespace RhubarbGeekNz.AtYourService
{
    internal class Program
    {
        static void Main(string[] args)
        {
            IHelloWorld helloWorld = Activator.CreateInstance(Type.GetTypeFromProgID("RhubarbGeekNz.AtYourService")) as IHelloWorld;

            string result = helloWorld.GetMessage(args.Length == 0 ? 1 : Int32.Parse(args[0]));

            Console.WriteLine($"{result}");
        }
    }
}
