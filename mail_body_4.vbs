Dim arrFileLines() 
Dim z
dim j
i = 0 
j = 0
c = 1
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set objFile = objFSO.OpenTextFile("C:\Users\IBM_ADMIN\Downloads\Stephen Marri\BaRclAyS\fLeX\pre_EOD_amol_beta_with_mail\Pre_EOD_REPORT.txt", 1) 
Set objoutputFile = objFSO.CreateTextFile("C:\Users\IBM_ADMIN\Downloads\Stephen Marri\BaRclAyS\fLeX\pre_EOD_amol_beta_with_mail\beta.txt")
Do Until objFile.AtEndOfStream 
 Redim Preserve arrFileLines(i) 
 arrFileLines(i) = objFile.ReadLine 
 i = i + 1 
Loop 
 


Dim ar
ar = Split("PRE-CHECK,Unauthorized_Contract,Unauthorized_Maintenance,Unauthorized_Batch,Unauthorized_TIL_VALULT,Unauthorized_Clearing,Unauthorized_Rates,Unauthorized_Liquidation,Unauthorized_Check_Book,Unauthorized_Check_Details,Unauthorized_Payments,Unauthorized_Amount_Blocks,Unauthorized_MIS_Adjustments,Unauthorized_TRANSACTION,Unauthorized_Messages,Unauthorized_CL_EVENTS,FCY_BALANCE_MISMATCH,LCY_BALANCE_MISMATCH,PC_Exceptions_Check,DEBUG_ON,EODM RUN STATUS,BRANCH_INFO,STTM_CUST_ACCOUNT,SMTB_CURRENT_USERS,STTM_BRANCH,STTM_DATES,STTM_AEOD_DATES", ",")


For l = Lbound(arrFileLines) to UBound(arrFileLines) Step 1 
a=split(arrFileLines(l))
for each x in a
for each y in ar
if x=y then
j=j+1
m=m & y & vbcrlf
end if
next 
if x="Attention!!!" then
z=z & c & ". " & arrFileLines(l) & vbcrlf
n=n & c & ". " & ar(j-1) & vbcrlf
c= c+1
end if
next
next
msgbox z
msgbox n

