using System;
using System.Runtime.InteropServices;

namespace DemoDocxApp
{
    static class Globals
    {

        // Used to load CradleAPI.dll and its dependant libraries
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Ansi)]
        extern static bool SetDllDirectory(string lpPathName);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Ansi)]
        extern static IntPtr LoadLibrary(string lpFileName);


        // Connection Settings
        public const string CRADLE_CDS_HOST = "SCO-SUITZ02";
        public const string CRADLE_HOME = @"C:\Program Files\Cradle\bin\exe\windows";
        public const string CRADLE_API = "CradleAPI.dll";
        public static string CRADLE_PROJECT_CODE = "";
        public const string CRADLE_USERNAME = "MANAGER";
        public const string CRADLE_PASSWORD = "MANAGER";
        public static string[] Args = new string[7];
        /// <summary>
        /// Loads CradleAPI.dll so we can run outside of the Cradle installation directory
        /// </summary>
        /// <remarks></remarks>
        public static void Load_CradleAPI()
        {

            String CRADLEHOME = System.Environment
                .GetEnvironmentVariable("CRADLEHOME", EnvironmentVariableTarget.Machine);
            String CRADLE_CDS_HOST = System.Environment
               .GetEnvironmentVariable("CRADLE_CDS_HOST", EnvironmentVariableTarget.Machine);
            String CRADLE_API_HOME = CRADLEHOME + @"\bin\exe\windows";

            if (SetDllDirectory(CRADLE_API_HOME))
            {
                IntPtr ptr = LoadLibrary(CRADLE_API);
                if ((ptr == IntPtr.Zero))
                    throw new DllNotFoundException("Could not find CradleAPI.dll");
            }
        }
        public static void GetArgs()
        {
            Args = Environment.CommandLine.Split('&');

            //Args = new string[] { "exe","TBL1", "ПД", "БЛ1", "БЛ2" };
            CRADLE_PROJECT_CODE = Globals.Args[1];
        }
    }
}
