// A class for parsing, storing, and outputting doses and volumes, specifically those that are calculated
// at a specific absolute/relative volume or dose.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Forms;
using System.IO;

namespace VMS.TPS.Common.Model.API
{
    class DVHQuery
    {
        // Private Fields
        private string query; // i.e. V%_90%
        private string inputMetric;
        private double inputValue;
        private string outputMetric;
        private double outputValue;
        private DoseValuePresentation doseOutputPresentation;
        private VolumePresentation volumeOutputPresentation;

        private PlanningItem plan;
        private Structure structure;

        // Used to identify the header in Excel when the
        // structure is null.
        private string nullStructName;
        private double ptvVolume;

        private DVHData dvhData;

        // For both Plan Setups and Plan Sums.
        private DoseValue? totalPrescribedDose;


        // Constructor
        public DVHQuery(string query,
                        PlanningItem plan,
                        Structure structure,
                        DoseValue? totalPrescribedDose = null,
                        string nullStructName = null,
                        double ptvVolume = 0,
                        DVHData dvhData = null)
        {
            this.query = query;
            this.inputMetric = "";
            this.inputValue = 0.0;
            this.outputMetric = "";
            this.outputValue = 0.0;

            this.plan = plan;
            this.structure = structure;

            if (plan is PlanSetup pSetup)
            {
                this.totalPrescribedDose = pSetup.TotalPrescribedDose;
            }
            else if (plan is PlanSum pSum)
            {
                this.totalPrescribedDose = totalPrescribedDose;
            }

            this.nullStructName = nullStructName;
            this.ptvVolume = ptvVolume;

            this.dvhData = dvhData;

            CalculateDVHAtPoint();
        }

        // Getters
        public string GetStructureName()
        {
            return nullStructName;
        }

        public double GetDVHValue()
        {
            return outputValue;
        }

        public string GetDVHMetric()
        {
            return outputMetric;
        }

        public string GetQueryString()
        {
            StringBuilder sb = new StringBuilder();

            if (query.Equals("dmax", StringComparison.OrdinalIgnoreCase))
            {
                return "Max Dose (cGy)";
            }
            else if (query.Equals("dmin", StringComparison.OrdinalIgnoreCase))
            {
                return "Min Dose (cGy)";
            }
            else if (query.Equals("dmean", StringComparison.OrdinalIgnoreCase))
            {
                return "Mean Dose (cGy)";
            }
            else if (query.Equals("p_dmax", StringComparison.OrdinalIgnoreCase))
            {
                return "Relative Max Dose";
            }
            else if (query.Equals("svol", StringComparison.OrdinalIgnoreCase))
            {
                string structName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(nullStructName.ToLower());
                string acronym = "";

                if (structName.ToUpper().Contains("PTV"))
                {
                    acronym = "PTV";
                }
                else if (structName.ToUpper().Contains("CTV"))
                {
                    acronym = "CTV";
                }
                else if (structName.ToUpper().Contains("GTV"))
                {
                    acronym = "GTV";
                }
                else if (structName.ToUpper().Contains("ITV"))
                {
                    acronym = "ITV";
                }

                if (!string.IsNullOrWhiteSpace(acronym))
                {
                    int substringStart = structName.ToUpper().IndexOf(acronym);
                    int substringEnd = substringStart + acronym.Length;

                    sb.Append(structName.Substring(0, substringStart));
                    sb.Append(structName.Substring(substringStart, acronym.Length).ToUpper());
                    sb.Append(structName.Substring(substringEnd, structName.Length - substringEnd));

                    structName = sb.ToString();
                }
                return $"{structName} Volume (cc)";
            }
            else if (query.Equals("hi", StringComparison.OrdinalIgnoreCase))
            {
                return "Homogeneity Index";
            }
            else if (query.StartsWith("ui", StringComparison.OrdinalIgnoreCase))
            {
                return "Conformity Index";
            }
            else if (query.StartsWith("ci", StringComparison.OrdinalIgnoreCase))
            {
                return "Uniformity Index";
            }
            else if (query.StartsWith("gi", StringComparison.OrdinalIgnoreCase))
            {
                return "Gradient Index";
            }
            else if (query.Equals("eqsd", StringComparison.OrdinalIgnoreCase))
            {
                return "Equivalent Sphere Diameter";
            }

            sb.Append(query.ToUpper()[0]);
            sb.Append(this.inputValue);
            sb.Append(this.inputMetric);

            if (!string.IsNullOrWhiteSpace(outputMetric))
            {
                sb.Append(string.Format(" ({0})", outputMetric));
            }
            return sb.ToString();
        }


        // Private Methods
        private void CalculateMaxDose()
        {
            // If DVHData is null, the structure exists but does not have DVH data.
            if (structure == null || this.dvhData == null)
            {
                this.outputValue = Double.NaN;
                return;
            }

            DoseValue dv = dvhData.MaxDose;
            this.outputMetric = dv.UnitAsString;
            this.outputValue = dv.Dose;
        }

        private void CalculateMinDose()
        {
            if (structure == null || this.dvhData == null)
            {
                this.outputValue = Double.NaN;
                return;
            }

            DoseValue dv = dvhData.MinDose;
            this.outputMetric = dv.UnitAsString;
            this.outputValue = dv.Dose;
        }

        private void CalculateMeanDose()
        {
            if (structure == null || this.dvhData == null)
            {
                this.outputValue = Double.NaN;
                return;
            }

            DoseValue dv = dvhData.MeanDose;
            this.outputMetric = dv.UnitAsString;
            this.outputValue = dv.Dose;
        }

        private void CalculateMaxDoseRatio()
        {
            if (structure == null || this.dvhData == null)
            {
                this.outputValue = Double.NaN;
                return;
            }

            DoseValue dv = dvhData.MaxDose;

            if (plan is PlanSetup pSetup)
            {
                this.outputMetric = dv.UnitAsString;

                if (pSetup.TotalPrescribedDose.Dose == 0 || double.IsNaN(pSetup.TotalPrescribedDose.Dose))
                {
                    this.outputValue = double.NaN;
                    return;
                }
                this.outputValue = dv.Dose * 100.0 / pSetup.TotalPrescribedDose.Dose;
            }
            else if (plan is PlanSum)
            {
                if (totalPrescribedDose is DoseValue planSumDose)
                {
                    this.outputMetric = dv.UnitAsString;

                    if (planSumDose.Dose == 0 || double.IsNaN(planSumDose.Dose))
                    {
                        this.outputValue = double.NaN;
                        return;
                    }
                    this.outputValue = dv.Dose * 100.0 / planSumDose.Dose;
                }
            }
        }

        private void CalculateStructureVolume()
        {
            this.outputMetric = "cc";
            if (structure == null)
            {
                this.outputValue = Double.NaN;
                return;
            }
            this.outputValue = structure.Volume;
        }

        private void CalculateHomogeneityIndex()
        {
            if (plan == null)
            {
                this.outputValue = Double.NaN;
                return;
            }

            DVHQuery dose2 = new DVHQuery("dcGy_2%", this.plan, this.structure, this.totalPrescribedDose,
                                              this.nullStructName, this.ptvVolume, this.dvhData);
            DVHQuery dose98 = new DVHQuery("dcGy_98%", this.plan, this.structure, this.totalPrescribedDose,
                                           this.nullStructName, this.ptvVolume, this.dvhData);

            if (plan is PlanSetup pSetup)
            {
                if (pSetup.TotalPrescribedDose.Dose == 0 || double.IsNaN(pSetup.TotalPrescribedDose.Dose))
                {
                    this.outputValue = double.NaN;
                    return;
                }
                this.outputValue = (dose2.GetDVHValue() - dose98.GetDVHValue()) / pSetup.TotalPrescribedDose.Dose;
            }
            else if (plan is PlanSum)
            {
                if (totalPrescribedDose is DoseValue planSumDose)
                {
                    if (planSumDose.Dose == 0 || double.IsNaN(planSumDose.Dose))
                    {
                        this.outputValue = double.NaN;
                        return;
                    }
                    this.outputValue = (dose2.GetDVHValue() - dose98.GetDVHValue()) / planSumDose.Dose;
                }
            }
        }

        private void CalculateUniformityIndex()
        {
            string dividend, divisor;

            if (query.Equals("ui", StringComparison.OrdinalIgnoreCase))
            {
                dividend = "dcGy_5%";
                divisor = "dcGy_95%";
            }
            else
            {
                dividend = query.Split('_').LastOrDefault().Split('/').FirstOrDefault();
                divisor = query.Split('_').LastOrDefault().Split('/').LastOrDefault();
                dividend = "dcGy_" + dividend + "%";
                divisor = "dcGy_" + divisor + "%";
            }

            DVHQuery dose5 = new DVHQuery(dividend, this.plan, this.structure, this.totalPrescribedDose,
                                          this.nullStructName, this.ptvVolume, this.dvhData);
            DVHQuery dose95 = new DVHQuery(divisor, this.plan, this.structure, this.totalPrescribedDose,
                                          this.nullStructName, this.ptvVolume, this.dvhData);

            if (dose95 == null || dose95.GetDVHValue() == 0 || double.IsNaN(dose95.GetDVHValue()))
            {
                this.outputValue = double.NaN;
                return;
            }
            this.outputValue = dose5.GetDVHValue() / dose95.GetDVHValue();
        }

        private void CalculateGradientIndex()
        {
            string dividend, divisor;

            if (query.Equals("gi", StringComparison.OrdinalIgnoreCase))
            {
                dividend = "dcGy_50%";
                divisor = "dcGy_100%";
            }
            else
            {
                dividend = query.Split('_').LastOrDefault().Split('/').FirstOrDefault();
                divisor = query.Split('_').LastOrDefault().Split('/').LastOrDefault();
                dividend = "dcGy_" + dividend + "%";
                divisor = "dcGy_" + divisor + "%";
            }

            DVHQuery dose50 = new DVHQuery(dividend, this.plan, this.structure, this.totalPrescribedDose,
                                          this.nullStructName, this.ptvVolume, this.dvhData);
            DVHQuery dose100 = new DVHQuery(divisor, this.plan, this.structure, this.totalPrescribedDose,
                                          this.nullStructName, this.ptvVolume, this.dvhData);

            if (dose100 == null || dose100.GetDVHValue() == 0 || double.IsNaN(dose100.GetDVHValue()))
            {
                this.outputValue = double.NaN;
                return;
            }
            this.outputValue = dose50.GetDVHValue() / dose100.GetDVHValue();
        }

        private void CalculateConformityIndex()
        {
            if (structure == null)
            {
                this.outputValue = Double.NaN;
                return;
            }

            string refIsodoseLine;

            if (query.Equals("ci", StringComparison.OrdinalIgnoreCase))
            {
                refIsodoseLine = "vcc_95%";
            }
            else
            {
                refIsodoseLine = query.Split('_').LastOrDefault();
                refIsodoseLine = "vcc_" + refIsodoseLine + "%";
            }

            DVHQuery volume95 = new DVHQuery(refIsodoseLine, this.plan, this.structure, this.totalPrescribedDose,
                                            this.nullStructName, this.ptvVolume, this.dvhData);
            double targetVolume = structure.Volume;

            if (targetVolume == 0 || double.IsNaN(targetVolume))
            {
                this.outputValue = double.NaN;
                return;
            }
            this.outputValue = volume95.GetDVHValue() / targetVolume;
        }

        private void CalculateEquivalentSphereDiameter()
        {
            if (structure == null)
            {
                this.outputValue = Double.NaN;
                return;
            }

            this.outputValue = (Math.Pow(((3.0 / (4.0 * Math.PI)) * structure.Volume), (1.0 / 3.0))) * 2.0;
        }

        public void CalculateDVHAtPoint()
        {
            //string inputInfo = @"\\dc3-pr-files\MedPhysics Backup\Coop Students\2019\Term 3 - Autumn\Quinton Tennant\Input Information.txt";
            uint timeout = 0; //messagebox timeout timer
			string tempIDLocInternal = @"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Info Transfer Internal.txt";
            string tempID = File.ReadLines(@tempIDLocInternal).Skip(1).Take(1).First();
			string inputInfo = File.ReadLines(@tempIDLocInternal).Skip(0).Take(1).First();
            string failedExtractionLoc = File.ReadLines(inputInfo).Skip(36).Take(1).First();

            // Dose and Volume presentation needs to be absolute for conversion and getting DVH.
            DoseValuePresentation dvp = DoseValuePresentation.Absolute;
            VolumePresentation vp = VolumePresentation.AbsoluteCm3;

            if (query.Equals("dmax", StringComparison.OrdinalIgnoreCase))
            {
                CalculateMaxDose();
                return;
            }
            else if (query.Equals("dmin", StringComparison.OrdinalIgnoreCase))
            {
                CalculateMinDose();
                return;
            }
            else if (query.Equals("dmean", StringComparison.OrdinalIgnoreCase))
            {
                CalculateMeanDose();
                return;
            }
            else if (query.Equals("p_dmax", StringComparison.OrdinalIgnoreCase))
            {
                CalculateMaxDoseRatio();
                return;
            }
            else if (query.Equals("svol", StringComparison.OrdinalIgnoreCase))
            {
                CalculateStructureVolume();
                return;
            }
            else if (query.Equals("hi", StringComparison.OrdinalIgnoreCase))
            {
                CalculateHomogeneityIndex();
                return;
            }
            else if (query.StartsWith("ui", StringComparison.OrdinalIgnoreCase))
            {
                CalculateUniformityIndex();
                return;
            }
            else if (query.StartsWith("gi", StringComparison.OrdinalIgnoreCase))
            {
                CalculateGradientIndex();
                return;
            }
            else if (query.StartsWith("ci", StringComparison.OrdinalIgnoreCase))
            {
                CalculateConformityIndex();
                return;
            }
            else if (query.Equals("eqsd", StringComparison.OrdinalIgnoreCase))
            {
                CalculateEquivalentSphereDiameter();
                return;
            }

            string x = query.Split('_').FirstOrDefault();
            string y = query.Split('_').LastOrDefault();

            if (x.StartsWith("V", StringComparison.OrdinalIgnoreCase) ||
                x.StartsWith("R", StringComparison.OrdinalIgnoreCase))
            {
                // The metric the volume will be outputted in.
                if (x.EndsWith("cc", StringComparison.OrdinalIgnoreCase))
                {
                    this.volumeOutputPresentation = VolumePresentation.AbsoluteCm3;
                    this.outputMetric = "cc";
                }
                else
                {
                    // Assume the default case is outputting the volume in % (relative).
                    this.volumeOutputPresentation = VolumePresentation.Relative;
                    this.outputMetric = "%";
                }

                // Used for converting dose metric.
                DoseValue.DoseUnit doseUnit;

                // Check what metric the dose was inputted in.
                if (y.EndsWith("cGy", StringComparison.OrdinalIgnoreCase))
                {
                    this.inputMetric = "cGy";
                    doseUnit = DoseValue.DoseUnit.cGy;
                    this.inputValue = Convert.ToDouble(y.Remove(y.Length - 3));
                }
                else if (y.EndsWith("Gy", StringComparison.OrdinalIgnoreCase))
                {
                    this.inputMetric = "Gy";
                    doseUnit = DoseValue.DoseUnit.Gy;
                    this.inputValue = Convert.ToDouble(y.Remove(y.Length - 2));
                }
                else if (y.EndsWith("%", StringComparison.OrdinalIgnoreCase))
                {
                    this.inputMetric = "%";
                    doseUnit = DoseValue.DoseUnit.Percent;
                    this.inputValue = Convert.ToDouble(y.Remove(y.Length - 1));
                }
                else
                {
                    // Assume the default case is that dose is inputted in % (relative).
                    this.inputMetric = "%";
                    doseUnit = DoseValue.DoseUnit.Percent;
                    this.inputValue = Convert.ToDouble(y);
                }

                if (structure == null || this.dvhData == null)
                {
                    this.outputValue = Double.NaN;
                    return;
                }

                // Convert the dose to cGy.
                DoseValue doseValue = new DoseValue(this.inputValue, doseUnit);

                // If plan sum, check if we have a valid prescribed dose for the plan sum.
                PlanSumPrescribedDoseErrorCheck();

                DoseValue doseValueInCgy = doseValue.ConvertDoseUnits(DoseValue.DoseUnit.cGy, this.totalPrescribedDose);

                double volumeInCc = DVHExtensions.VolumeAtDose(dvhData, doseValueInCgy.Dose);

                if (x.StartsWith("R", StringComparison.OrdinalIgnoreCase))
                {
                    // Convert volume from cc (absolute) to % (relative).
                    double volumeInPercent = volumeInCc.ConvertVolumeUnits("cc", VolumePresentation.Relative, dvhData);

                    if (double.IsNaN(ptvVolume) || ptvVolume == 0.0)
                    {
                        //Self-exiting messagebox pop up (change timeout timer to change exit time)
                        MessageBoxEx.Show($"The PTV could not be found for plan {plan.Id}.", "Warning",
                                            MessageBoxButtons.OK, MessageBoxIcon.Warning, timeout);
                        using (System.IO.StreamWriter failedExtractionFile =
                            new System.IO.StreamWriter(failedExtractionLoc, true))
                        {
                            failedExtractionFile.Write(tempID);
                            failedExtractionFile.WriteLine($" -- The PTV could not be found for plan {plan.Id}.");
                        }

                        String[] TextFileLines = File.ReadAllLines(failedExtractionLoc);
                        String[] TextFileLinesDist;
                        TextFileLinesDist = TextFileLines.Distinct().ToArray();
                        File.WriteAllLines(failedExtractionLoc, TextFileLinesDist);

                        this.outputValue = Double.NaN;
                        return;
                    }

                    this.outputMetric = "";
                    this.outputValue = volumeInPercent * structure.Volume / (100 * ptvVolume);

                    return;
                }

                // Get the DVH with the metric of the volume we want outputted.
                this.outputValue = volumeInCc.ConvertVolumeUnits("cc", volumeOutputPresentation, dvhData);
            }
            else if (x.StartsWith("D", StringComparison.OrdinalIgnoreCase))
            {
                // The metric the dose will be outputted in.
                if (x.EndsWith("%", StringComparison.OrdinalIgnoreCase))
                {
                    this.doseOutputPresentation = DoseValuePresentation.Relative;
                    this.outputMetric = "%";
                }
                else if (x.EndsWith("cGy", StringComparison.OrdinalIgnoreCase))
                {
                    this.doseOutputPresentation = DoseValuePresentation.Absolute;
                    this.outputMetric = "cGy";
                }
                else if (x.EndsWith("Gy", StringComparison.OrdinalIgnoreCase))
                {
                    this.doseOutputPresentation = DoseValuePresentation.Absolute;
                    this.outputMetric = "Gy";
                }
                else
                {
                    // Assume the default case is outputting the dose in cGy (absolute).
                    this.doseOutputPresentation = DoseValuePresentation.Absolute;
                    this.outputMetric = "cGy";
                }

                // Used for converting volume metric.
                VolumePresentation volumeUnit;

                // Check what metric the volume was inputted in.
                if (y.EndsWith("%", StringComparison.OrdinalIgnoreCase))
                {
                    this.inputMetric = "%";
                    volumeUnit = VolumePresentation.Relative;
                    this.inputValue = Convert.ToDouble(y.Remove(y.Length - 1));
                }
                else if (y.EndsWith("cc", StringComparison.OrdinalIgnoreCase))
                {
                    this.inputMetric = "cc";
                    volumeUnit = VolumePresentation.AbsoluteCm3;
                    this.inputValue = Convert.ToDouble(y.Remove(y.Length - 2));
                }
                else
                {
                    // Assume the default case is that volume is inputted in cc (absolute).
                    this.inputMetric = "cc";
                    volumeUnit = VolumePresentation.AbsoluteCm3;
                    this.inputValue = Convert.ToDouble(y);
                }

                if (structure == null || this.dvhData == null)
                {
                    this.outputValue = Double.NaN;
                    return;
                }

                // Convert the volume to cc.
                double volume = this.inputValue;
                double volumeInCc = volume.ConvertVolumeUnits(this.inputMetric, vp, dvhData);

                DoseValue doseValueInCgy = DVHExtensions.DoseAtVolume(dvhData, volumeInCc);

                // Get the dose with the metric we want outputted.
                if (outputMetric.Equals("cGy", StringComparison.OrdinalIgnoreCase))
                {
                    this.outputValue = doseValueInCgy.ConvertDoseUnits(DoseValue.DoseUnit.cGy, this.totalPrescribedDose).Dose;
                }
                else if (outputMetric.Equals("Gy", StringComparison.OrdinalIgnoreCase))
                {
                    this.outputValue = doseValueInCgy.ConvertDoseUnits(DoseValue.DoseUnit.Gy, this.totalPrescribedDose).Dose;
                }
                else if (outputMetric.Equals("%", StringComparison.OrdinalIgnoreCase))
                {
                    PlanSumPrescribedDoseErrorCheck();
                    this.outputValue = doseValueInCgy.ConvertDoseUnits(DoseValue.DoseUnit.Percent, this.totalPrescribedDose).Dose;
                }
            }
            else
            {
                this.outputValue = Double.NaN;
            }
        }

        public void PlanSumPrescribedDoseErrorCheck()
        {
            //string inputInfo = @"\\dc3-pr-files\MedPhysics Backup\Coop Students\2019\Term 3 - Autumn\Quinton Tennant\Input Information.txt";
            uint timeout = 0; //messagebox timeout timer
            string tempIDLocInternal = @"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\Info Transfer Internal.txt";
            string tempID = File.ReadLines(@tempIDLocInternal).Skip(1).Take(1).First();
            string inputInfo = File.ReadLines(@tempIDLocInternal).Skip(0).Take(1).First();
            string failedExtractionLoc = File.ReadLines(inputInfo).Skip(36).Take(1).First();

            if (plan is PlanSum)
            {
                if (totalPrescribedDose is null)
                {
                    //Self-exiting messagebox pop up (change timeout timer to change exit time)
                    MessageBoxEx.Show("The prescribed dose for the plan sum could not be found. Please check if " +
                    "it was inputted correctly in the template.", "Error",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error, timeout);
                    using (System.IO.StreamWriter failedExtractionFile =
                        new System.IO.StreamWriter(failedExtractionLoc, true))
                    {
                        failedExtractionFile.Write(tempID);
                        failedExtractionFile.WriteLine($" -- The prescribed dose for the plan sum could not be found. Please check if it was inputted correctly in the template.");
                    }
                    String[] TextFileLines = File.ReadAllLines(failedExtractionLoc);
                    String[] TextFileLinesDist;
                    TextFileLinesDist = TextFileLines.Distinct().ToArray();
                    File.WriteAllLines(failedExtractionLoc, TextFileLinesDist);
                }
            }
        }
    }
}