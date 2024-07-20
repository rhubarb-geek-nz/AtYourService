/***********************************
 * Copyright (c) 2024 Roger Brown.
 * Licensed under the MIT License.
 ****/

using System;
using System.Reflection;

namespace dispnet
{
    internal class Program
    {
        static void Main(string[] args)
        {
            object helloWorld = Activator.CreateInstance(Type.GetTypeFromProgID("RhubarbGeekNz.AtYourService"));

            int hint = args.Length == 0 ? 1 : Int32.Parse(args[0]);

            object result = helloWorld.GetType().InvokeMember("GetMessage", BindingFlags.Public | BindingFlags.Instance | BindingFlags.InvokeMethod, null, helloWorld, new object[] { hint });

            Console.WriteLine($"{result}");
        }
    }
}
