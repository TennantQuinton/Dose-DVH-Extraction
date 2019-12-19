using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SQL_Script_Calling
{
	class Program
	{
		static void Main(string[] args)
		{
			string sqlConnectionString = ("Server=grc505n; Database=variansystem; User Id=reports; Password=reports;");

			// Create the connection to the resource!
			// This is the connection, that is established and
			// will be available throughout this block.
			using (SqlConnection conn = new SqlConnection())
			{
				// Create the connectionString
				// Trusted_Connection is used to denote the connection uses Windows Authentication
				conn.ConnectionString = sqlConnectionString; //"Server=grc505n; Database=variansystem; User Id = reports; Password = reports";
				conn.Open();
				// Create the command
				SqlCommand command = new SqlCommand(@"
select distinct
(p.PatientId) PatientId,
(c.CourseId) CourseId,
(ps.PlanSetupId) Plan_Name,
dvh.Structures,
(p.DateOfBirth) Date_Of_Birth,
(vc.VolumeCode) Body_Region,
rtp.PrescribedDose,
rtp.NoFractions,
--mlcp.MLCPlanType,
--f.GantryRtnDirection,
(ps.CreationDate) Plan_CreationDate

from PlanSetup ps
join Course c on c.CourseSer = ps.CourseSer
inner join Patient p on p.PatientSer = c.PatientSer
inner join RTPlan rtp on rtp.PlanSetupSer = ps.PlanSetupSer
inner join DoseContribution dc
	on rtp.RTPlanSer = dc.RTPlanSer
inner join RefPoint rp
	on dc.RefPointSer = rp.RefPointSer
inner join PatientVolume pv
	on rp.PatientVolumeSer = pv.PatientVolumeSer
inner join VolumeCode vc
	on pv.VolumeCodeSer = vc.VolumeCodeSer
inner join Radiation r
	on ps.PlanSetupSer = r.PlanSetupSer
inner join ExternalFieldCommon fc
	on r.RadiationSer = fc.RadiationSer
--inner join MLCPlan mlcp
--	on r.RadiationSer = mlcp.RadiationSer
inner join DVH dvh
	on rtp.PlanSetupSer = dvh.PlanSetupSer

where 1=1
and cast(ps.CreationDate as date) between '01/01/2003' and '12/30/2019'
and vc.VolumeCode in ('')
and ps.PlanSetupId like '%SRS%'
and p.PatientId not like '%$%'
--and c.CourseId not like ('%QA%')
--and f.GantryRtnDirection = 'CW'
--and f.GantryRtnDirection = 'NONE'
--and mlcp.MLCPlanType = 'StdMLCPlan'
--and mlcp.MLCPlanType = 'DynMLCPlan'
and c.CourseId in ('1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20')
--and rtp.PrescribedDose > '10'
--and rtp.NoFractions = '23'
order by p.PatientId"
, conn);
				// Add the parameters.
				command.Parameters.Add(new SqlParameter("0", 1));

				/* Get the rows and display on the screen! 
                 * This section of the code has the basic code
                 * that will display the content from the Database Table
                 * on the screen using an SqlDataReader. */

				Microsoft.Office.Interop.Excel.Application oXL;
				Microsoft.Office.Interop.Excel._Workbook oWB;
				Microsoft.Office.Interop.Excel._Worksheet oSheet;
				Microsoft.Office.Interop.Excel.Range oRng;
				
				object misvalue = System.Reflection.Missing.Value;
				//Start Excel and get Application object.
				oXL = new Microsoft.Office.Interop.Excel.Application();
				oXL.DisplayAlerts = false;
				oXL.Visible = true;

				//Get a new workbook.
				oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
				oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

				using (SqlDataReader reader = command.ExecuteReader())
				{
					//Add table headers going cell by cell.
					oSheet.Cells[1, 1] = "PatientId";
					oSheet.Cells[1, 2] = "CourseId";
					oSheet.Cells[1, 3] = "Plan_Name";
					oSheet.Cells[1, 4] = "Structures";
					//Console.WriteLine("PatientId \t | CourseId \t | Plan_Name \t | Structures\t");
					int count = 2;
					while (reader.Read())
					{
						oSheet.Cells[count, 1] = $"{reader[0]}";
						oSheet.Cells[count, 2] = $"{reader[1]}";
						oSheet.Cells[count, 3] = $"{reader[2]}";
						oSheet.Cells[count, 4] = $"{reader[3]}";

						//Console.WriteLine(String.Format($"{reader[0]} \t | {reader[1]} \t | {reader[2]} \t | {reader[3]}"));
						count++;
					}
				}

				oXL.Visible = true;
				oXL.UserControl = true;
				
				oWB.SaveAs($@"\\dc3-pr-files\MedPhysics Backup\Data Extractions\Automation\SQL Database Spreadsheets\Brain\test.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
					true, true, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
					Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

				//oWB.Close();
				//oXL.Quit();
			}
		}
	}
}
