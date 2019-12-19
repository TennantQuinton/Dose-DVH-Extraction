// A wrapper class for a list of DVHQueries collected by plan.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VMS.TPS.Common.Model.Types;

namespace VMS.TPS.Common.Model.API
{
    class PlanQueries
    {
        // Private Fields
        private List<DVHQuery> queriesByPlan;

        private Patient patient;
        private Course course;
        private PlanningItem plan;

        private bool beamDataInitialized;
        private List<Beam> vmatList;
        private List<Beam> doseDynamicList;
        private List<Beam> staticList;
        private List<Beam> coneList;
        private List<Beam> electronList;
        private double? ssd1;
        private double? ssd2;

        private HashSet<string> energyList;

        // Plan Sum specific fields
        private List<PlanQueries> planSumPlans;
        private DoseValue? planSumPrescribedDose;


        // Constructor
        public PlanQueries(Patient patient,
                    Course course,
                    PlanningItem plan,
                    DoseValue? planSumPrescribedDose)
        {
            this.queriesByPlan = new List<DVHQuery>();

            this.patient = patient;
            this.course = course;
            this.plan = plan;

            this.beamDataInitialized = false;
            this.vmatList = new List<Beam>();
            this.doseDynamicList = new List<Beam>();
            this.staticList = new List<Beam>();
            this.coneList = new List<Beam>();
            this.electronList = new List<Beam>();
            this.ssd1 = null;
            this.ssd2 = null;

            this.energyList = new HashSet<string>();

            this.planSumPlans = new List<PlanQueries>();
            this.planSumPrescribedDose = planSumPrescribedDose;
        }
        
        // Accessors
        public List<DVHQuery> GetDVHQueryList()
        {
            return queriesByPlan;
        }

        public void AddDVHQuery(DVHQuery dvhQuery)
        {
            queriesByPlan.Add(dvhQuery);
        }

        public int GetDVHQueryCount()
        {
            return queriesByPlan.Count;
        }

        public List<PlanQueries> GetPlanSumPlans()
        {
            return planSumPlans;
        }

        public void AddPlanSumPlan(PlanQueries planQueries)
        {
            planSumPlans.Add(planQueries);
        }

        // Methods
        public string GetPatientID()
        {
            if (patient == null)
            {
                return "";
            }
            return patient.Id;
        }

        public string GetCourseID()
        {
            if (course == null)
            {
                return "";
            }
            return course.Id;
        }

        public string GetPlanID()
        {
            if (plan == null)
            {
                return "";
            }
            return plan.Id;
        }

        public DateTime? GetPlanCreationDateTime()
        {
            if (plan == null)
            {
                return null;
            }

            DateTime? planCreationDate = null;

            if (plan is PlanSetup pSetup)
            {
                if (plan.CreationDateTime != null)
                {
                    planCreationDate = pSetup.CreationDateTime;
                }
            }
            else if (plan is PlanSum pSum)
            {
                if (plan.HistoryDateTime != null)
                {
                    planCreationDate = pSum.HistoryDateTime;
                }
                else if (plan.CreationDateTime != null)
                {
                    planCreationDate = pSum.CreationDateTime;
                }
            }
            return planCreationDate;
        }

        // Calculates the patient's age when the plan was created.
        public int? GetAgeAtPlanCreationDate()
        {
            if (plan == null || patient == null)
            {
                return null;
            }

            DateTime? birthDate = patient.DateOfBirth;
            DateTime? planCreationDate = GetPlanCreationDateTime();

            if (birthDate == null || planCreationDate == null)
            {
                return null;
            }

            int age = ((DateTime)planCreationDate).Year - ((DateTime)birthDate).Year;
            if (((DateTime)birthDate).Date > ((DateTime)planCreationDate).AddYears(-age))
            {
                age--;
            }
            return age;
        }

        public double? GetDoseFx()
        {
            if (plan == null)
            {
                return null;
            }

            if (plan is PlanSetup pSetup)
            {
                try
                {
                    DoseValue dosePerFraction = pSetup.UniqueFractionation.PrescribedDosePerFraction;
                    return dosePerFraction.ConvertDoseUnits(DoseValue.DoseUnit.Gy, pSetup.TotalPrescribedDose).Dose;
                }
                catch (Exception)
                {
                    return null;
                }
            }
            else if (plan is PlanSum)
            {
                return null;
            }
            return null;
        }

        public int? GetNumOfFractions()
        {
            if (plan == null)
            {
                return null;
            }

            if (plan is PlanSetup pSetup)
            {
                try
                {
                    return pSetup.UniqueFractionation.NumberOfFractions;
                }
                catch (Exception)
                {
                    return null;
                }
            }
            else if (plan is PlanSum)
            {
                return null;
            }
            return null;
        }

        public double? GetTotalDoseInGy()
        {
            if (plan == null)
            {
                return null;
            }

            if (plan is PlanSetup pSetup)
            {
                return pSetup.TotalPrescribedDose.ConvertDoseUnits(DoseValue.DoseUnit.Gy, pSetup.TotalPrescribedDose).Dose;
            }
            else if (plan is PlanSum)
            {
                if (planSumPrescribedDose is DoseValue dv)
                {
                    return dv.ConvertDoseUnits(DoseValue.DoseUnit.Gy, planSumPrescribedDose).Dose;
                }
                else if (planSumPrescribedDose is null)
                {
                    return null;
                }
            }
            return null;
        }

        public double? GetNumOfBeams()
        {
            if (plan == null)
            {
                return null;
            }

            int count = 0;
            if (plan is PlanSetup pSetup)
            {
                foreach (Beam beam in pSetup.Beams)
                {
                    if (!beam.IsSetupField)
                    {
                        count++;
                    }
                }
                return count;
            }
            else if (plan is PlanSum)
            {
                return null;
            }
            return null;
        }

        private void InitializeBeamData()
        {
            if (plan is PlanSetup pSetup)
            {
                // Sort beams into lists by Plan type.
                foreach (Beam beam in pSetup.Beams)
                {
                    if (!beam.IsSetupField)
                    {
                        // MLC beam.
                        if (beam.MLCPlanType == MLCPlanType.VMAT)
                        {
                            vmatList.Add(beam);
                            energyList.Add(beam.EnergyModeDisplayName);
                        }
                        else if (beam.MLCPlanType == MLCPlanType.DoseDynamic)
                        {
                            doseDynamicList.Add(beam);
                            energyList.Add(beam.EnergyModeDisplayName);
                        }
                        else if (beam.MLCPlanType == MLCPlanType.Static)
                        {
                            staticList.Add(beam);
                            energyList.Add(beam.EnergyModeDisplayName);
                        }
                        else if (beam.Applicator != null)
                        {
                            // Cone beam.
                            if (beam.Applicator.Id.ToUpper().EndsWith("CC"))
                            {
                                coneList.Add(beam);
                                energyList.Add(beam.EnergyModeDisplayName);
                            }
                            // Electron beam.
                            if (beam.Applicator.Id.ToUpper().StartsWith("A"))
                            {
                                electronList.Add(beam);
                                energyList.Add(beam.EnergyModeDisplayName);
                            }
                        }
                    }
                }

                // Initialize the appropriate SSDs based on the overriding plan type.
                // 1 or more VMAT beams means it's a VMAT.
                if (vmatList.Count >= 1)
                {
                    // SSDs should be blank for VMAT.
                    this.ssd1 = null;
                    this.ssd2 = null;
                }
                // If no VMAT, 1 or more Dose Dynamic beams means Dose Dynamic.
                else if (doseDynamicList.Count >= 1)
                {
                    this.ssd1 = doseDynamicList[0].SSD / 10; ;

                    if (doseDynamicList.Count >= 2)
                    {
                        this.ssd2 = doseDynamicList[1].SSD / 10; ;
                    }
                }
                // If no VMAT and no Dose Dynamic, it's a Static.
                else if (staticList.Count >= 1)
                {
                    this.ssd1 = staticList[0].SSD / 10; ;

                    if (staticList.Count >= 2)
                    {
                        this.ssd2 = staticList[1].SSD / 10; ;
                    }
                }
                else if (coneList.Count >= 1)
                {
                    // SSDs should be blank for cone.
                    this.ssd1 = null;
                    this.ssd2 = null;
                }
                else if (electronList.Count >= 1)
                {
                    // Electron usually only has one SSD.
                    this.ssd1 = electronList[0].SSD / 10; ;

                    if (electronList.Count >= 2)
                    {
                        this.ssd2 = electronList[1].SSD / 10; ;
                    }
                }
            }
            else if (plan is PlanSum)
            {

            }
            this.beamDataInitialized = true;
        }

        public string GetMlcPlanType()
        {
            if (plan == null)
            {
                return null;
            }

            if (this.beamDataInitialized == false)
            {
                InitializeBeamData();
            }

            if (plan is PlanSetup)
            {
                if (vmatList.Count >= 1)
                {
                    return "VMAT";
                }
                else if (doseDynamicList.Count >= 1)
                {
                    return "Dose Dynamic";
                }
                else if (staticList.Count >= 1)
                {
                    return "Static";
                }
                else if (coneList.Count >= 1)
                {
                    return "Cone";
                }
                else if (electronList.Count >= 1)
                {
                    return "Electron";
                }
                else
                {
                    return "";
                }
            }
            else if (plan is PlanSum)
            {
                return "";
            }
            return "";
        }

        public double? GetSsd1()
        {
            if (plan == null)
            {
                return null;
            }

            if (this.beamDataInitialized == false)
            {
                InitializeBeamData();
            }

            if (plan is PlanSetup)
            {
                if (this.ssd1 != null)
                {
                    return this.ssd1;
                }
            }
            else if (plan is PlanSum)
            {
                return null;
            }
            return null;
        }

        public double? GetSsd2()
        {
            if (plan == null)
            {
                return null;
            }

            if (this.beamDataInitialized == false)
            {
                InitializeBeamData();
            }

            if (plan is PlanSetup)
            {
                if (this.ssd2 != null)
                {
                    return this.ssd2;
                }
            }
            else if (plan is PlanSum)
            {
                return null;
            }
            return null;
        }

        public double? GetSsdDivision()
        {
            if (plan == null)
            {
                return null;
            }

            if (this.beamDataInitialized == false)
            {
                InitializeBeamData();
            }

            if (plan is PlanSetup)
            {
                if (GetSsd1() != null && GetSsd2() != null)
                {
                    return (100 - GetSsd1()) + (100 - GetSsd2());
                }
            }
            else if (plan is PlanSum)
            {
                return null;
            }
            return null;
        }

        public string GetFieldEnergies()
        {
            if (plan == null)
            {
                return null;
            }

            if (this.beamDataInitialized == false)
            {
                InitializeBeamData();
            }

            if (energyList.Count == 0)
            {
                return null;
            }

            if (plan is PlanSetup)
            {
                StringBuilder sb = new StringBuilder();
                bool first = true;

                foreach (string energy in energyList)
                {
                    if (!first)
                    {
                        sb.Append(", ");
                    }
                    sb.Append(energy);
                    first = false;
                }
                return sb.ToString();
            }
            else if (plan is PlanSum)
            {
                return null;
            }
            return null;
        }

        public bool IsPlanSum()
        {
            if (plan is PlanSum)
            {
                return true;
            }
            return false;
        }
    }
}
