using System;
using System.Runtime.InteropServices;

namespace CostService
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
        public static string CRADLE_PROJECT_CODE = "TBL1";
        public const string CRADLE_USERNAME = "MANAGER";
        public const string CRADLE_PASSWORD = "MANAGER";
        public static string[] Args;



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
            CRADLE_PROJECT_CODE = Args[1];
            //Args = new string[] { "exe","TBL1", "ПД", "БЛ1", "БЛ2" };
        }
        public static string[] GetQuerys()
        {
            // Проверка фазы проекта
            switch (Args[2])
            {
                case "АФК":
                    return new string[] { "Требование ТЗ АФК - ALL", "Раздел ТЗ АФК - ALL" };
                case "ПД":
                    return new string[] { "Требование ТЗ ПД - ALL", "Раздел ТЗ ПД - ALL" };
                case "РД":
                    return new string[] { "Требование ТЗ РД - ALL", "Раздел ТЗ РД - ALL" };
                case "СМР":
                    return new string[] { "Требование ТЗ СМР - ALL", "Раздел ТЗ СМР - ALL" };
                default:
                    return null;
            }

        }
    }
}