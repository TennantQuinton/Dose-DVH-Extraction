using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Clear_This_Folder
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Last Updated: 2019-12-17 by Tennant, Quinton");

			string workingLoc = @"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Info Transfer External.txt";
			string property = "";

			if (File.ReadAllLines(workingLoc).Contains("propertyDose"))
			{
				property = "Dose";
			}
			else if (File.ReadAllLines(workingLoc).Contains("propertyDVH"))
			{
				property = "DVH";
			}
			else
			{
				MessageBox.Show("Dose or DVH not selected", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}

			string inputInfo = $@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\{property}\Input Information Text Files\General {property} Input Information.txt";
            string archiveLoc = File.ReadLines(inputInfo).Skip(20).Take(1).First();
            string path = $@"{archiveLoc}\";

			DateTime localDate = DateTime.Now;
            string localDateString = localDate.ToString("yyyy-MM-dd");

			if (File.Exists(inputInfo))
			{
				if (Directory.Exists(path))
				{
					// Delete all files in a directory    
					string[] folders = Directory.GetDirectories(path);
					foreach (string folder in folders)
					{
						Directory.Delete(folder, true);
						Console.WriteLine($"'{Path.GetFileName(folder)}' Deleted.");
					}
				}
				else
				{
					MessageBox.Show($"Path {path} does not exist", "Pathing Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
			else
			{
				MessageBox.Show($"No Input Information file available for {property}.", "Pathing Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}

			Random rnd = new Random();
            int num = rnd.Next(1, 1000);
            if (num > 0 && num < 25)
            {
                Console.WriteLine("Folders have been EXTERMINATED!");
                Console.WriteLine(@"
                  ___
 -@@~-    D>=G==='   '.
                |======|
                |======|
            )--/]IIIIII]
               |_______|
               C O O O D
              C O  O  O D
             C  O  O  O  D
             C__O__O__O__D
            [_____________]");
            }
            else if (num > 26 && num < 226)
            {
                Console.WriteLine("Folders have been removed.");
            }
            else if (num > 227 && num < 229)
            {
                Console.WriteLine("I'm sorry Dave, I'm afraid I can't do that.");
            }
            else if (num > 230 && num < 245)
            {
                Console.WriteLine(@"
                                \            .  ./
                             \   .: ;' .::, .:..   /
                               (M^^.^~       ~:.' ).
                         -   (/  .           . . \ \)  -
  O                         ((| :. ~        ^  :. .|))
 |\\                     -   (\- |  \       /  |  /)  -
 |  T                         -\  \            /  /-
/ \[_]..........................\  \The Files/  /");
            }
            else if (num > 246 && num < 250)
            {
                Console.WriteLine(@"
   ..-^~~~^-..
 .~           ~.
(;:           :;)
 (:           :)
   ':._   _.:'
       | |
     (=====)
       | |
       | |
       | |
   ((/files\))");
            }
            else if (num > 995 && num < 997)
            {
                Console.WriteLine(@"Directory Cleared!");
                Console.WriteLine(@" 
You Found an Easter Egg!

 ,adPPYba,  ,adPPYb,d8  ,adPPYb,d8  
a8P_____88 a8/    `Y88 a8/    `Y88  
8PP``````` 8b       88 8b       88  
`8b,   ,aa `8a,   ,d88 `8a,   ,d88  
 `'Ybbd8'`  `'YbbdP'Y8  `'YbbdP'Y8  
            aa,    ,88  aa,    ,88  
             `Y8bbdP`    `Y8bbdP`");
            }
            else if (num > 997 && num < 999)
            {
                Console.WriteLine("Directory Cleared!");
                Console.WriteLine(@"
You Foun an Easter Egg!
         .--.
       .'    ',
     .'        ',
    /\/\/\/\/\/\/\
   /\/\/\/\/\/\/\/\
  Y                Y
  |/\/\/\/\/\/\/\/\|
  |\/\/\/\/\/\/\/\/|
  Y                Y
   \/\/\/\/\/\/\/\/
    \/\/\/\/\/\/\/
     `.        .'
       `..__..'    ");
            }
            else
            {
                Console.WriteLine("Directory Emptied!");
            }
            Console.Write("Press any key to exit...");
            Console.ReadLine();
        }
    }
}
