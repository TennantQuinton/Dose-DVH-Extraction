using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Clean_Slate_Protocol
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Last Updated: 2019-11-28 by Tennant, Quinton");
            Console.WriteLine("");

			//deleting patient ID lists DOSE
			string[] LocationDose = Directory.GetFiles(@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Dose\Patient ID Lists", "*.txt");
			foreach (string item in LocationDose)
			{
				if (item.Contains("Dose Patient ID List.txt"))
				{
					File.Delete(item);
					Console.WriteLine($"{item} deleted.");
				}
			}
            string templatePathDose = @"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Dose\Template Creation Output";
			//deleting templates DOSE
			if (Directory.Exists(templatePathDose))
            {
                string[] foldersDose = Directory.GetDirectories(templatePathDose);
                foreach (var folderDose in foldersDose)
                {
                    string[] filesDose = Directory.GetFiles(folderDose);
                    foreach (var fileDose in filesDose)
                    {
                        File.Delete(fileDose);
                    }

                    Directory.Delete(folderDose);
                }
                Console.WriteLine($"{templatePathDose} emptied.");
            }
			//deleting failed extractions DOSE
			string[] failedDose = Directory.GetFiles(@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Dose", "*.txt");
			foreach (string item in failedDose)
			{
				if (item.Contains("Dose Failed Extractions.txt"))
				{
					File.Delete(item);
					Console.WriteLine($"{item} deleted.");
				}
			}



			//deleting patient ID lists DVH
			string[] LocationDVH = Directory.GetFiles(@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\DVH\Patient ID Lists", "*.txt");
			foreach (string item in LocationDVH)
			{
				if (item.Contains("DVH Patient ID List.txt"))
				{
					File.Delete(item);
					Console.WriteLine($"{item} deleted.");
				}
			}
			string templatePathDVH = @"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\DVH\Template Creation Output";
			//deleting templates DVH
			if (Directory.Exists(templatePathDVH))
			{
				string[] foldersDVH = Directory.GetDirectories(templatePathDVH);
				foreach (var folderDVH in foldersDVH)
				{
					string[] filesDVH = Directory.GetFiles(folderDVH);
					foreach (var fileDVH in filesDVH)
					{
						File.Delete(fileDVH);
					}

					Directory.Delete(folderDVH);
				}
				Console.WriteLine($"{templatePathDVH} emptied.");
			}
			//deleting failed extractions DVH
			string[] failedDVH = Directory.GetFiles(@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\DVH", "*.txt");
			foreach (string item in failedDVH)
			{
				if (item.Contains("DVH Failed Extractions.txt"))
				{
					File.Delete(item);
					Console.WriteLine($"{item} deleted.");
				}
			}

			Console.WriteLine("Done! ");
            Console.Write("Press any key to exit...");
            Console.ReadLine();
        }
    }
}
