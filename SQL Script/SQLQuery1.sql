select distinct
(p.PatientId) PatientId,
(p.DateOfBirth) Date_Of_Birth,
(c.CourseId) CourseId,
(vc.VolumeCode) Body_Region,
(ps.PlanSetupId) Plan_Name,
rtp.PrescribedDose,
rtp.NoFractions,
mlcp.MLCPlanType,
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
inner join MLCPlan mlcp
	on r.RadiationSer = mlcp.RadiationSer

where 1=1
and cast(ps.CreationDate as date) between '01/01/2003' and '12/30/2019'
--and vc.VolumeCode in ('BRAI')
and ps.PlanSetupId like '%BRAI%'
--and c.CourseId not like ('%QA%')
--and f.GantryRtnDirection = 'CW'
--and f.GantryRtnDirection = 'NONE'
and mlcp.MLCPlanType = 'StdMLCPlan'
--and mlcp.MLCPlanType = 'DynMLCPlan'
and c.CourseId in ('1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20')
and rtp.PrescribedDose > '10'
--and rtp.NoFractions = '23'
order by p.PatientId
;