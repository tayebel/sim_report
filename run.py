#open RESInsinght 
#import the reservoir simulation case (both summary file and EGRID file) into RESInsinght. 
from sim_report import SimReport
s=input("Would you like to create datbase tables into your Ms SQL server (YES or NO ) ")
if s=="YES":
    ss="ACTIVE"
else:
    ss=None
unit=input("Enter the unit that you are using for your simulation (write FIELD or METRIC ): ")
cn=int(input("check the case number in ResInsight and enter it (Ex: 1): "))
c=input("Enter the Current Operator : ")
o=input("Enter the Original Operator : ")
Rl=float(input("Enter the Rlat (Ex: 29): "))
Rlo=float(input("Enter the Rlong (Ex: -95.43): "))
M=input("MeasuredDepth, Optional (Ex: 5000): ")
T=input("TrueVerticalDepth, Optional (Ex: 5000): ")
sim=SimReport(unit=unit)

#generate the report
sim.sheet_report(case_num=cn,Rlat=Rl,Rlong=Rlo,MeasuredDepth=M,TrueVerticalDepth=T,CurrentOperator=c,OriginalOperator=o,sql_t=ss)