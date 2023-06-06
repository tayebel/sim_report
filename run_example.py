
from sim_report import SimReport
sim=SimReport(unit="METRIC")
#you can add where you would like to save the Excel report
saving_path = "Res_sim_report.xlsx"
#check the case number in ResIn
case_num=1
Rlat=29.731329
Rlong=-95.427296
#optional
MeasuredDepth=5000
TrueVerticalDepth=5000
#generate the report
sim.sheet_report(case_num,Rlat,Rlong,saving_path,MeasuredDepth,TrueVerticalDepth)