'**********************************************************************************************************************
'Test Name:		创建群 (ec登陆后主面板)
'Purpose:			
'Requirement:	
'Note:			 
'Starting Page:	创建群
'Created by:		zhangyang	 			
'Creation date:	29/06/2018
'
'Modification History: 
'Date:				Changed by:			Purpose:
'**********************************************************************************************************************
Option Explicit

'定义参数变量
Dim strType, strSubType, strNotes
Dim strGroupName, strGrpupMemberName1

'==================================================================
'依据业务脚本调用，确定执行步骤
Call SetActionTableRow(Parameter("Action"))
'==================================================================

strType = DataTable("Type", dtLocalSheet)
strSubType = DataTable("SubType", dtLocalSheet)
strNotes = DataTable("Notes", dtLocalSheet)

strGroupName = EvaluateInputParam(DataTable("群名称", dtLocalSheet))
strGroupNotes = EvaluateInputParam(DataTable("群公告", dtLocalSheet))
strGrpupMemberName1 = EvaluateInputParam(DataTable("EC号码", dtLocalSheet))


Call CloseOptionalDialog(2)

'*****************************************************************
'脚本说明：检查当前是否在登录界面
'*****************************************************************
If Window("EC10登录界面_操作项").InsightObject("登录按钮").Exist(5) Then
	Reporter.ReportEvent micPass, "Verify Page - 登录界面","At 登录界面 Page"
Else
	Reporter.ReportEvent micFail, "Verify Page - 登录界面","Not at 登录界面 Page"
	Call ExitRun()
End If
'*****************************************************************
'脚本说明：登录界面可操作步骤
'*****************************************************************
Select Case Lcase(strType)
	Case Lcase("创建群")
		Select Case Lcase(strSubType)
			'1.输入群名&群公告
			Case Lcase("输入群名")
				Window("创建群").InsightObject("群名输入框").type strGroupName
                Window("创建群").InsightObject("群公告输入框").type strGroupNotes

			'2.选择只有管理员才能添加群成员
			Case Lcase("选择只有管理员才能添加群成员")
		  		Window("创建群").InsightObject("选择只有管理员才能添加成员").Click

			'3.选择任何人都能添加群成员
			Case Lcase("选择任何人都能添加群成员")
				Window("创建群").InsightObject("选择任何人都能添加群成员").Click

			'4.单击下一步
			Case Lcase("点击下一步")
				Window("创建群").InsightObject("下一步").Click
			 

			'5.选择群成员，创建完成
			Case Lcase("选择群成员完成创建")
				
				Do until Window("创建群").InsightObject("添加群成员名").Exist(5)
					wait 1
				Loop 
				Window("创建群").InsightObject("添加群成员名").type strGrpupMemberName1
				  wait 1
				  Window("创建群").DblClick 95,180,micLeftBtn 
				  wait 1
				  Window("创建群").InsightObject("群创建完成按钮").Click

		End Select
End Select