using System.Reflection;
using System.Media;

namespace Schedule_Calculator_Pro
{
    static class Consts
    {
        public static string AssemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
        public static bool SaveLoadInProgress = false;
        public static SoundPlayer JobDoneSound = new SoundPlayer(Consts.LocalToGlobal("\\Sounds\\jobdonenotif.wav"));

        public static void Init()
        {
            JobDoneSound.LoadAsync();
        }

        public static string LocalToGlobal(string path) // convert local path (relative to the .EXE) to global path
        {
            return AssemblyPath.Replace("\\Schedule Calculator Pro.exe", path);
        }
    }
}