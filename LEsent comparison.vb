'Snooze Button - LE
If Macro.GetField("CX.LE.VALIDATION.CHANGES") <> "Yes" then
  Button12.enabled = False
 Else Button12.enabled = True
End If

'Snooze Button - CD 
 If Macro.GetField("CX.CD.VALIDATION.CHANGES") <> "Yes" then
  Button13.enabled = False
 Else Button13.enabled = True
End If

Dim cdSent As String
cdSent = Loan.Fields("3977").Value

Dim lEsent As String
leSent = Loan.Field("3153").Value

'LE loan Comparison
If Not (leSent.IsEmpty()) and 
	'If Macro.GetField("4000")<>Macro.GetField("CX.LE.4000") Then
	' TextBox161.BackColor=Color.Red
	'End If
	'If Macro.GetField("4002")<>Macro.GetField("CX.LE.4002") Then
	' TextBox162.BackColor=Color.Red
	'End If
	'If Macro.GetField("4004")<>Macro.GetField("CX.LE.4004") Then
	' TextBox165.BackColor=Color.Red
	'End If
	'If Macro.GetField("4006")<>Macro.GetField("CX.LE.4006") Then
	' TextBox166.BackColor=Color.Red
	'End If
	'If Macro.GetField("910")<>Macro.GetField("CX.LE.910") Then
	' TextBox163.BackColor=Color.Red
	'End If
	'If Macro.GetField("911")<>Macro.GetField("CX.LE.911") Then
	' TextBox167.BackColor=Color.Red
	'End If
	If Macro.GetField("MORNET.X67")<>Macro.GetField("CX.LE.MORNET.X67") Then
	 TextBox183.BackColor=Color.Red
	End If
	If Macro.GetField("4")<>Macro.GetField("CX.LE.4") Then
	 TextBox168.BackColor=Color.Red
	End If
	If Macro.GetField("136")<>Macro.GetField("CX.LE.136") Then
	 TextBox169.BackColor=Color.Red
	End If
	If Macro.GetField("2")<>Macro.GetField("CX.LE.2") Then
	 TextBox170.BackColor=Color.Red
	End If
	' If Macro.GetField("353")<>Macro.GetField("CX.LE.353") Then
	 ' TextBox171.BackColor=Color.Red
	' End If
	' If Macro.GetField("976")<>Macro.GetField("CX.LE.976") Then
	 ' TextBox176.BackColor=Color.Red
	' End If
	' If Macro.GetField("740")<>Macro.GetField("CX.LE.740") Then
	 ' TextBox178.BackColor=Color.Red
	' End If
	If Macro.GetField("742")<>Macro.GetField("CX.LE.742") Then
	 TextBox177.BackColor=Color.Red
	End If
	If Macro.GetField("19")<>Macro.GetField("CX.LE.19") Then
	 TextBox173.BackColor=Color.Red 
	End If
	If Macro.GetField("1172")<>Macro.GetField("CX.LE.1172") Then
	 TextBox172.BackColor=Color.Red 
	End If
	If Macro.GetField("1401")<>Macro.GetField("CX.LE.1401") Then
	 TextBox174.BackColor=Color.Red
	End If
	If Macro.GetField("2293")<>Macro.GetField("CX.LE.2293") Then
	 TextBox175.BackColor=Color.Red
	End If
	'If Macro.GetField("VASUMM.X23")<>Macro.GetField("CX.LE.VASUMM.X23") Then
	' TextBox179.BackColor=Color.Red
	'End If
	If Macro.GetField("356")<>Macro.GetField("CX.LE.356") Then
	 TextBox185.BackColor=Color.Red
	End If
	' If Macro.GetField("11")<>Macro.GetField("CX.LE.11") Then
	 ' TextBox186.BackColor=Color.Red
	' End If
	' If Macro.GetField("16")<>Macro.GetField("CX.LE.16") Then
	 ' TextBox187.BackColor=Color.Red
	' End If
	'If Macro.GetField("1041")<>Macro.GetField("CX.LE.1041") Then
	' TextBox189.BackColor=Color.Red
	'End If
	'If Macro.GetField("1811")<>Macro.GetField("CX.LE.1811") Then
	' TextBox188.BackColor=Color.Red
	'End If
	If Macro.GetField("3")<>Macro.GetField("CX.LE.3") Then
	 TextBox180.BackColor=Color.Red
	End If
	If Macro.GetField("799")<>Macro.GetField("CX.LE.799") Then
	 TextBox62.BackColor=Color.Red
	End If
	' **** CLIENT ADDED FIELDS ****
	If Macro.GetField("2400")<>Macro.GetField("CX.LE.2400") Then
	 TextBox27.BackColor=Color.Red
	End If
	If Macro.GetField("762")<>Macro.GetField("CX.LE.762") Then
	 TextBox2.BackColor=Color.Red
	End If
	If Macro.GetField("LE1.X5")<>Macro.GetField("CX.LE.LE1.X5") Then
	 TextBox68.BackColor=Color.Red
	End If
	If Macro.GetField("675")<>Macro.GetField("CX.LE.675") Then
	 TextBox71.BackColor=Color.Red
	End If
	' If Macro.GetField("2312")<>Macro.GetField("CX.LE.2312") Then
	 ' TextBox181.BackColor=Color.Red
	' End If
	' If Macro.GetField("12")<>Macro.GetField("CX.LE.12") Then
	 ' TextBox190.BackColor=Color.Red
	' End If
	' If Macro.GetField("14")<>Macro.GetField("CX.LE.14") Then
	 ' TextBox191.BackColor=Color.Red
	' End If
	'If Macro.GetField("1012")<>Macro.GetField("CX.LE.1012") Then
	'TextBox192.BackColor=Color.Red
	'End If
	If Macro.GetField("428")<>Macro.GetField("CX.LE.428") Then
	 TextBox182.BackColor=Color.Red
	End If
	' If Macro.GetField("608")<>Macro.GetField("CX.LE.608") Then
	 ' Checkbox9.BackColor=Color.Red
	' End If
	 If Macro.GetField("1045")<>Macro.GetField("CX.LE.1045") Then
	  TextBox184.BackColor=Color.Red
	 End If
	' If Macro.GetField("688")<>Macro.GetField("CX.LE.688") Then
	 ' TextBox194.BackColor=Color.Red
	' End If
	' If Macro.GetField("689")<>Macro.GetField("CX.LE.689") Then
	 ' TextBox195.BackColor=Color.Red
	' End If
	' If Macro.GetField("1959")<>Macro.GetField("CX.LE.1959") Then
	 ' TextBox193.BackColor=Color.Red
	' End If
	' If Macro.GetField("230")<>Macro.GetField("CX.LE.230") Then
	 ' TextBox200.BackColor=Color.Red
	' End If 
	' If Macro.GetField("1405")<>Macro.GetField("CX.LE.1405") Then
	 ' TextBox199.BackColor=Color.Red 
	' End If
	' If Macro.GetField("233")<>Macro.GetField("CX.LE.233") Then
	 ' TextBox198.BackColor=Color.Red 
	' End If
	' If Macro.GetField("NEWHUD2.X23")<>Macro.GetField("CX.LE.NEWHUD2.X23") Then
	 ' TextBox10.BackColor=Color.Red 
	' End If
	' If Macro.GetField("BORPCC")<>Macro.GetField("CX.LE.BORPCC") Then
	 ' TextBox11.BackColor=Color.Red 
	' End If
	' If Macro.GetField("143")<>Macro.GetField("CX.LE.143") Then
	 ' TextBox12.BackColor=Color.Red 
	' End If
	' If Macro.GetField("BKRPCC")<>Macro.GetField("CX.LE.BKRPCC") Then
	 ' TextBox4.BackColor=Color.Red 
	' End If
	' If Macro.GetField("LENPCC")<>Macro.GetField("CX.LE.LENPCC") Then
	 ' TextBox5.BackColor=Color.Red 
	' End If
	' If Macro.GetField("OTHPCC")<>Macro.GetField("CX.LE.OTHPCC") Then
	 ' TextBox6.BackColor=Color.Red 
	' End If
	' If Macro.GetField("TNBPCC")<>Macro.GetField("CX.LE.TNBPCC") Then
	 ' TextBox7.BackColor=Color.Red 
	' End If
	' If Macro.GetField("NEWHUD.X1149")<>Macro.GetField("CX.LE.NEWHUD.X1149") Then
	 ' TextBox9.BackColor=Color.Red 
	' End If
	' If Macro.GetField("TOTPCC")<>Macro.GetField("CX.LE.TOTPCC") Then
	 ' TextBox8.BackColor=Color.Red 
	' End If
End If



'CD loan Comparison
 
'If cdSent <> "" and Macro.GetField("4000")<>Macro.GetField("CX.CD.4000") Then
' TextBox237.BackColor=Color.Red
'End If
'If cdSent <> "" and Macro.GetField("4002")<>Macro.GetField("CX.CD.4002") Then
' TextBox238.BackColor=Color.Red
'End If
'If cdSent <> "" and Macro.GetField("4004")<>Macro.GetField("CX.CD.4004") Then
' TextBox234.BackColor=Color.Red
'End If
'If cdSent <> "" and Macro.GetField("4006")<>Macro.GetField("CX.CD.4006") Then
' TextBox235.BackColor=Color.Red
'End If
'If cdSent <> "" and Macro.GetField("910")<>Macro.GetField("CX.CD.910") Then
' TextBox239.BackColor=Color.Red
'End If
'If cdSent <> "" and Macro.GetField("911")<>Macro.GetField("CX.CD.911") Then
' TextBox236.BackColor=Color.Red
'End If
If cdSent <> "" and Macro.GetField("MORNET.X67")<>Macro.GetField("CX.CD.MORNET.X67") Then
 TextBox224.BackColor=Color.Red
End If
If cdSent <> "" and Macro.GetField("4")<>Macro.GetField("CX.CD.4") Then
 TextBox209.BackColor=Color.Red
End If
If cdSent <> "" and Macro.GetField("136")<>Macro.GetField("CX.CD.136") Then
 TextBox210.BackColor=Color.Red
End If
If cdSent <> "" and Macro.GetField("2")<>Macro.GetField("CX.CD.2") Then
 TextBox211.BackColor=Color.Red
End If
' If cdSent <> "" and Macro.GetField("353")<>Macro.GetField("CX.CD.353") Then
 ' TextBox212.BackColor=Color.Red
' End If
' If cdSent <> "" and Macro.GetField("976")<>Macro.GetField("CX.CD.976") Then
 ' TextBox217.BackColor=Color.Red
' End If
' If cdSent <> "" and Macro.GetField("740")<>Macro.GetField("CX.CD.740") Then
 ' TextBox219.BackColor=Color.Red
' End If
' If cdSent <> "" and Macro.GetField("742")<>Macro.GetField("CX.CD.742") Then
 ' TextBox218.BackColor=Color.Red
' End If
If cdSent <> "" and Macro.GetField("19")<>Macro.GetField("CX.CD.19") Then
 TextBox214.BackColor=Color.Red
End If
If cdSent <> "" and Macro.GetField("1172")<>Macro.GetField("CX.CD.1172") Then
 TextBox213.BackColor=Color.Red
End If
If cdSent <> "" and Macro.GetField("1401")<>Macro.GetField("CX.CD.1401") Then
 TextBox215.BackColor=Color.Red
End If
If cdSent <> "" and Macro.GetField("2293")<>Macro.GetField("CX.CD.2293") Then
 TextBox216.BackColor=Color.Red
End If
'If cdSent <> "" and Macro.GetField("VASUMM.X23")<>Macro.GetField("CX.CD.VASUMM.X23") Then
' TextBox220.BackColor=Color.Red
'End If
If cdSent <> "" and Macro.GetField("356")<>Macro.GetField("CX.CD.356") Then
 TextBox201.BackColor=Color.Red
End If
' If cdSent <> "" and Macro.GetField("11")<>Macro.GetField("CX.CD.11") Then
 ' TextBox202.BackColor=Color.Red
' End If
' If cdSent <> "" and Macro.GetField("16")<>Macro.GetField("CX.CD.16") Then
 ' TextBox203.BackColor=Color.Red
' End If
'If cdSent <> "" and Macro.GetField("1041")<>Macro.GetField("CX.CD.1041") Then
' TextBox205.BackColor=Color.Red
'End If
'If cdSent <> "" and Macro.GetField("1811")<>Macro.GetField("CX.CD.1811") Then
' TextBox204.BackColor=Color.Red
'End If
If cdSent <> "" and Macro.GetField("3")<>Macro.GetField("CX.CD.3") Then
 TextBox221.BackColor=Color.Red
End If
If cdSent <> "" and Macro.GetField("799")<>Macro.GetField("CX.CD.799") Then
 TextBox63.BackColor=Color.Red
End If
' *** CLIENT ADDED FIELDS ***
If cdSent <> "" and Macro.GetField("2400")<>Macro.GetField("CX.CD.2400") Then
 TextBox54.BackColor=Color.Red
End If
If cdSent <> "" and Macro.GetField("762")<>Macro.GetField("CX.CD.762") Then
 TextBox3.BackColor=Color.Red
End If
If cdSent <> "" and Macro.GetField("LE1.X5")<>Macro.GetField("CX.CD.LE1.X5") Then
 TextBox69.BackColor=Color.Red
End If
If cdSent <> "" and Macro.GetField("675")<>Macro.GetField("CX.CD.675") Then
 TextBox72.BackColor=Color.Red
End If
' If cdSent <> "" and Macro.GetField("2312")<>Macro.GetField("CX.CD.2312") Then
 ' TextBox222.BackColor=Color.Red
' End If
' If cdSent <> "" and Macro.GetField("12")<>Macro.GetField("CX.CD.12") Then
 ' TextBox206.BackColor=Color.Red
' End If
' If cdSent <> "" and Macro.GetField("14")<>Macro.GetField("CX.CD.14") Then
 ' TextBox207.BackColor=Color.Red
' End If
' If cdSent <> "" and Macro.GetField("1012")<>Macro.GetField("CX.CD.1012") Then
 ' TextBox208.BackColor=Color.Red
' End If
If cdSent <> "" and Macro.GetField("428")<>Macro.GetField("CX.CD.428") Then
 TextBox223.BackColor=Color.Red
End If
' If cdSent <> "" and Macro.GetField("608")<>Macro.GetField("CX.CD.608") Then
 ' Checkbox11.BackColor=Color.Red
' End If
 If cdSent <> "" and Macro.GetField("1045")<>Macro.GetField("CX.CD.1045") Then
  TextBox225.BackColor=Color.Red
 End If
' If cdSent <> "" and Macro.GetField("688")<>Macro.GetField("CX.CD.688") Then
 ' TextBox227.BackColor=Color.Red
' End If
' If cdSent <> "" and Macro.GetField("689")<>Macro.GetField("CX.CD.689") Then
 ' TextBox228.BackColor=Color.Red
' End If
' If cdSent <> "" and Macro.GetField("1959")<>Macro.GetField("CX.CD.1959") Then
 ' TextBox226.BackColor=Color.Red
' End If
' If cdSent <> "" and Macro.GetField("230")<>Macro.GetField("CX.CD.230") Then
 ' TextBox233.BackColor=Color.Red
' End If cdSent <> "" and 
' If cdSent <> "" and Macro.GetField("1405")<>Macro.GetField("CX.CD.1405") Then
 ' TextBox232.BackColor=Color.Red
' End If
' If cdSent <> "" and Macro.GetField("233")<>Macro.GetField("CX.CD.233") Then
 ' TextBox231.BackColor=Color.Red
' End If
' If cdSent <> "" and Macro.GetField("NEWHUD2.X23")<>Macro.GetField("CX.CD.NEWHUD2.X23") Then
 ' TextBox38.BackColor=Color.Red
' End If
' If cdSent <> "" and Macro.GetField("BORPCC")<>Macro.GetField("CX.CD.BORPCC") Then
 ' TextBox39.BackColor=Color.Red
' End If
' If cdSent <> "" and Macro.GetField("143")<>Macro.GetField("CX.CD.143") Then
 ' TextBox41.BackColor=Color.Red
' End If
' If cdSent <> "" and Macro.GetField("BKRPCC")<>Macro.GetField("CX.CD.BKRPCC") Then
 ' TextBox30.BackColor=Color.Red
' End If
' If cdSent <> "" and Macro.GetField("LENPCC")<>Macro.GetField("CX.CD.LENPCC") Then
 ' TextBox33.BackColor=Color.Red 
' End If
' If cdSent <> "" and Macro.GetField("OTHPCC")<>Macro.GetField("CX.CD.OTHPCC") Then
 ' TextBox34.BackColor=Color.Red 
' End If
' If cdSent <> "" and Macro.GetField("TNBPCC")<>Macro.GetField("CX.CD.TNBPCC") Then
 ' TextBox35.BackColor=Color.Red 
' End If
' If cdSent <> "" and Macro.GetField("NEWHUD.X1149")<>Macro.GetField("CX.CD.NEWHUD.X1149") Then
 ' TextBox37.BackColor=Color.Red
' End If
' If cdSent <> "" and Macro.GetField("TOTPCC")<>Macro.GetField("CX.CD.TOTPCC") Then
 ' TextBox36.BackColor=Color.Red
' End If