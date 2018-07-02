'**********************************************************************************************************************
'Test Name:		初始化
'Purpose:			
'Requirement:	
'Note:			 
'Starting Page:	
'Created by:		zhangyang	 			
'Creation date:	2017.10.19
'
'Modification History: 
'Date:				Changed by:			Purpose:
'**********************************************************************************************************************
Option Explicit

'Declaration and Initialization
Dim strType, strSubType, strNotes
Dim strFilePath
Dim i

'==================================================================
'Use the action keyword to find the correct row to run, if exists
Call SetActionTableRow(Parameter("Action"))
'==================================================================
strType = DataTable("Type", dtLocalSheet)
strSubType = DataTable("SubType", dtLocalSheet)
strNotes = DataTable("Notes", dtLocalSheet)

strFilePath = EvaluateInputParam(DataTable("启动文件路径", dtLocalSheet))				'启动文件路径

Call CloseOptionalDialog(2)
'==================================
'perform based on subtype
'==================================
Select Case Lcase(strType)
	Case Lcase("初始化进程")
		Select Case Lcase(strSubType)
			Case Lcase("清理EC进程")		'1、清理进程
				For  i = 1 To 2 Step 1
				systemutil.CloseProcessByName "EC.exe"
				systemutil.CloseProcessByName "ECBrowserPV.exe"
				systemutil.CloseProcessByName "ECBrowserSub.exe"
				systemutil.CloseProcessByName "ECHub.exe"
				systemutil.CloseProcessByName "ECPhone.exe"
				systemutil.CloseProcessByName "ECReportProc.exe"
				wait 1
			   	Next
			Case Lcase("启动EC进程")		'2、启动EC10
				systemutil.Run strFilePath
				If Window("EC10.0").InsightObject("登录按钮").Exist(5)  Then
				Reporter.ReportEvent micPass,"启动是否成功","启动成功！"	
				Else
                Reporter.ReportEvent micFail,"启动是否成功","启动失败！"	
                Call ExitRun()
				End If
			Case Lcase("保留")		'3、
				
		End Select
End Select