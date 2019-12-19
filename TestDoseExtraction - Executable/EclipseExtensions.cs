using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VMS.TPS.Common.Model.Types;
using StringExtensions;
using System.IO;

namespace VMS.TPS.Common.Model.API
{
    static class EclipseExtensions
    {
        // Generates a patient's study ID using the default conventions.
        public static string GenerateStudyID(this Patient patient)
        {
            if (patient == null)
            {
                return "";
            }

            string lastName = patient.LastName;
            string patientID = patient.Id;

            lastName = lastName.GetSubstringByLength(0, 3);
            lastName = lastName.CharFiller('_', 3);
            patientID = patientID.GetSubstringByLength(2, 4);
            patientID = patientID.CharFiller('_', 4);

            return (lastName + patientID).ToUpper();
        }

        // Checks if two strings are the same (case-insensitive).
        public static bool CheckIDMatch(string s1, string s2)
        {
            bool result = s1.Equals(s2, StringComparison.OrdinalIgnoreCase);
            return result;
        }

        public static string FormatCitrixSaveDirectory(string saveDirectory)
        {
            // Check whether we are running from Citrix.
            if (string.Equals(Environment.GetEnvironmentVariable("SESSIONNAME").Substring(0, 3),
                "ICA", StringComparison.OrdinalIgnoreCase))
            {
                Uri uri = new Uri(saveDirectory);
                // If the save directory is not specified to be the location of a network resource.
                if (!uri.IsUnc)
                {
                    return @"\\Client\" + saveDirectory.Replace(':', '$');
                }
            }
            return saveDirectory;
        }

        // Parses the prescribed dose from a string to a DoseValue.
        public static DoseValue ParsePrescribedDose(string dose)
        {
            double inputValue = Double.NaN;
            // Used for converting dose metric.
            DoseValue.DoseUnit doseUnit;

            // Check what metric the dose was inputted in.
            if (dose.EndsWith("cGy", StringComparison.OrdinalIgnoreCase))
            {
                doseUnit = DoseValue.DoseUnit.cGy;
                inputValue = Convert.ToDouble(dose.Remove(dose.Length - 3));
            }
            else if (dose.EndsWith("Gy", StringComparison.OrdinalIgnoreCase))
            {
                doseUnit = DoseValue.DoseUnit.Gy;
                inputValue = Convert.ToDouble(dose.Remove(dose.Length - 2));
            }
            else
            {
                // Assume the default case is that dose is inputted in cGy (absolute).
                doseUnit = DoseValue.DoseUnit.cGy;
                inputValue = Convert.ToDouble(dose);
            }
            return new DoseValue(inputValue, doseUnit);
        }

        // Converts the prescribed dose to cGy.
        // The prescribed dose must be in Gy or cGy (absolute).
        private static DoseValue GetPrescibedDoseInCgy(this DoseValue totalPrescribedDose)
        {
            if (totalPrescribedDose.Unit == DoseValue.DoseUnit.Gy)
            {
                // Make sure the total prescribed dose is in cGy.
                return totalPrescribedDose * 100.0;
            }
            else if (totalPrescribedDose.Unit == DoseValue.DoseUnit.Percent ||
                     totalPrescribedDose.Unit == DoseValue.DoseUnit.Unknown)
            {
                throw new ArgumentException("The total prescribed dose must be in Gy or cGy (absolute).");
            }
            return totalPrescribedDose;
        }

        // Converts the dose units.
        // The total prescribed dose must be passed as a parameter when coverting between absolute and relative and vice versa.
        public static DoseValue ConvertDoseUnits(this DoseValue inputDoseValue, DoseValue.DoseUnit outputUnit, DoseValue? totalPrescribedDose = null)
        {
            double value = inputDoseValue.Dose;
            DoseValue.DoseUnit inputUnit = inputDoseValue.Unit;

            // If the dose we want to convert is in cGy.
            if (inputUnit == DoseValue.DoseUnit.cGy)
            {
                if (outputUnit == DoseValue.DoseUnit.Gy)
                {
                    value = value / 100.0;
                }
                else if (outputUnit == DoseValue.DoseUnit.Percent)
                {
                    if (totalPrescribedDose is DoseValue dv)
                    {
                        dv = dv.GetPrescibedDoseInCgy();
                        value = (value * 100.0) / dv.Dose;
                    }
                    else if (totalPrescribedDose is null)
                    {
                        throw new ArgumentNullException("totalPrescribedDose is null; required for conversions between absolute and relative.");
                    }
                }
            }
            // If the dose we want to convert is in Gy.
            else if (inputUnit == DoseValue.DoseUnit.Gy)
            {
                if (outputUnit == DoseValue.DoseUnit.cGy)
                {
                    value = value * 100.0;
                }
                else if (outputUnit == DoseValue.DoseUnit.Percent)
                {
                    if (totalPrescribedDose is DoseValue dv)
                    {
                        dv = dv.GetPrescibedDoseInCgy();
                        value = ((value * 100.0) * 100.0) / dv.Dose;
                    }
                    else if (totalPrescribedDose is null)
                    {
                        throw new ArgumentNullException("totalPrescribedDose is null; required for conversions between absolute and relative.");
                    }
                }
            }
            // If the dose we want to convert is in percentage.
            else if (inputUnit == DoseValue.DoseUnit.Percent)
            {
                if (totalPrescribedDose is DoseValue dv)
                {
                    dv = dv.GetPrescibedDoseInCgy();
                    if (outputUnit == DoseValue.DoseUnit.cGy)
                    {
                        value = (value * dv.Dose) / 100.0;
                    }
                    else if (outputUnit == DoseValue.DoseUnit.Gy)
                    {
                        value = ((value * dv.Dose) / 100.0) / 100.0;
                    }
                }
                else if (totalPrescribedDose is null)
                {
                    throw new ArgumentNullException("totalPrescribedDose is null; required for conversions between absolute and relative.");
                }
            }
            // If the dose units are unknown.
            else
            {
                string tempIDLocInternal = @"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Info Transfer Internal.txt";
                string tempID = File.ReadLines(@tempIDLocInternal).Skip(1).Take(1).First();
                string inputInfo = File.ReadLines(@tempIDLocInternal).Skip(0).Take(1).First();
                string failedExtractionLoc = File.ReadLines(inputInfo).Skip(36).Take(1).First();

                //if extraction fails writes to failed extractions w/ ID and reason
                using (System.IO.StreamWriter failedExtractionFile =
                    new System.IO.StreamWriter(failedExtractionLoc, true))
                {
                    failedExtractionFile.Write(tempID);
                    failedExtractionFile.WriteLine(" -- The dose could not be converted because the units are unknown.");
                }

                String[] TextFileLines = File.ReadAllLines(failedExtractionLoc);
                String[] TextFileLinesDist;
                TextFileLinesDist = TextFileLines.Distinct().ToArray();
                File.WriteAllLines(failedExtractionLoc, TextFileLinesDist);
                //throw new ArgumentException("The dose could not be converted because the units are unknown.");
            }
            return new DoseValue(value, outputUnit);
        }

        // Converts the volume units.
        public static double ConvertVolumeUnits(this double inputVolume, string inputUnit, VolumePresentation outputUnit, DVHData dvhData)
        {
            double volumeAtZero = TPS.DVHExtensions.VolumeAtDose(dvhData, 0);
            if (inputUnit == "%")
            {
                if (outputUnit == VolumePresentation.AbsoluteCm3)
                {
                    // Convert volume from % (relative) to cc (absolute).
                    return inputVolume * volumeAtZero / 100;
                }
            }
            else if (inputUnit == "cc")
            {
                if (outputUnit == VolumePresentation.Relative)
                {
                    // Convert volume from cc (absolute) to % (relative).
                    return 100 * inputVolume / volumeAtZero;
                }
            }
            return inputVolume;
        }

        public static Course GetCourseByID(Patient patient, string courseID)
        {
            if (patient == null || patient.Courses == null)
            {
                return null;
            }
            return patient.Courses.FirstOrDefault(x => CheckIDMatch(x.Id, courseID));
        }

        public static PlanSetup GetPlanSetupByID(Course course, string planID)
        {
            if (course == null || course.PlanSetups == null)
            {
                return null;
            }
            return course.PlanSetups.FirstOrDefault(x => CheckIDMatch(x.Id, planID));
        }

        public static PlanSum GetPlanSumByID(Course course, string planID)
        {
            if (course == null || course.PlanSums == null)
            {
                return null;
            }
            return course.PlanSums.FirstOrDefault(x => CheckIDMatch(x.Id, planID));
        }

        public static Structure GetStructureByID(PlanningItem plan, string structID)
        {
            if (plan == null)
            {
                return null;
            }

            if (plan is PlanSetup pSetup)
            {
                if (pSetup.StructureSet == null|| pSetup.StructureSet.Structures == null)
                {
                    return null;
                }
                return pSetup.StructureSet.Structures.FirstOrDefault(x => EclipseExtensions.CheckIDMatch(x.Id, structID));
            }
            else if (plan is PlanSum pSum)
            {
                if (pSum.StructureSet == null || pSum.StructureSet.Structures == null)
                {
                    return null;
                }
                return pSum.StructureSet.Structures.FirstOrDefault(x => EclipseExtensions.CheckIDMatch(x.Id, structID));
            }
            return null;
        }
    }
}
