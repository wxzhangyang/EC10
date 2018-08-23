'**********************************************************************************************************************
'Test Name:		EC_10创建讨论组
'Purpose:			
'Requirement:	
'Note:			 
'Starting Page:	主界面
'Created by:		zhangyang	 			
'Creation date:	08/08/2018
'
'Modification History: 
'Date:				Changed by:			Purpose:
'**********************************************************************************************************************
Option Explicit

'定义参数变量
Dim strType, strSubType, strNotes
Dim strDiscussionsName, strDiscussionsMemberName1,strDiscussionsNotes

'==================================================================
'依据业务脚本调用，确定执行步骤
Call SetActionTableRow(Parameter("Action"))
'==================================================================

strType = DataTable("Type", dtLocalSheet)
strSubType = DataTable("SubType", dtLocalSheet)
strNotes = DataTable("Notes", dtLocalSheet)

strDiscussionsName = EvaluateInputParam(DataTable("讨论组名称", dtLocalSheet))
strDiscussionsMemberName1 = EvaluateInputParam(DataTable("EC号码", dtLocalSheet))

Call CloseOptionalDialog(2)

'*****************************************************************
'脚本说明：检查当前是否在登录界面
'*****************************************************************
'If Window("EC10登录界面_操作项").InsightObject("登录按钮").Exist(5) Then
'	Reporter.ReportEvent micPass, "Verify Page - 登录界面","At 登录界面 Page"
'Else
'	Reporter.ReportEvent micFail, "Verify Page - 登录界面","Not at 登录界面 Page"
'	Call ExitRun()
'End If
'*****************************************************************
'脚本说明：登录界面可操作步骤
'*****************************************************************
Select Case Lcase(strType)
	Case Lcase("讨论组")
		Select Case Lcase(strSubType)
			'1.选择添加讨论组成员
			Case Lcase("创建讨论组")
			Window("讨论组管理").DblClick 360,260,micLeftBtn 
			Window("讨论组管理").InsightObject("讨论组搜索").type strDiscussionsMemberName1
			 wait 1
			Window("讨论组管理").DblClick 120,106,micLeftBtn 
			Window("讨论组管理").InsightObject("确定").Click
			Window("EC10.0").InsightObject("输入框").type strDiscussionsMemberName1
			Window("EC10.0").InsightObject("消息发送").Click
			
			'2.更改讨论组名
			Case Lcase("更改讨论组名")
				Window("EC10.0").DblClick 390,30,micLeftBtn 
                window("EC10.0").InsightObject("讨论组名").type strDiscussionsName
                
			'3.退出讨论组
			Case Lcase("退出讨论组")
				Window("EC10.0").InsightObject("搜索讨论组").type strDiscussionsName
				wait 1
				Dim wshShell
                Set wshShell = CreateObject("Wscript.Shell")
                wshShell.SendKeys "{ENTER}" 
                Window("EC10.0").InsightObject("退出讨论组").Click

                Window("EC10.0").InsightObject("确定").Click

                Window("EC10.0").InsightObject("搜索讨论组").type strDiscussionsName
                   If Window("EC10.0").InsightObject("搜索结果校验").Exist(10) Then
                    Reporter.ReportEvent micPass, "是否正常退出讨论组?","退出成功!"
                Else
                    Reporter.ReportEvent micFail, "是否正常退出讨论组?","退出失败!"
                    Call ExitRun()
                End If
                wait 1
                Window("EC10.0").InsightObject("清除搜索框").Click

		End Select
End Select