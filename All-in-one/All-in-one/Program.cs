using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace All_in_one
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Last Updated: 2019-10-29 by Tennant, Quinton");

            var processA = Process.Start(@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Executables\1. Patient Template Creation.lnk");
            processA.WaitForExit();

            var processB = Process.Start(@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Executables\2. Extract Dose.lnk");
            processB.WaitForExit();
        }
    }
}
