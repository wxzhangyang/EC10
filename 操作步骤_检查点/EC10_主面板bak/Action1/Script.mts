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


Call CloseOptionalDialog(2)

'===================================================================                     判断是否处于登录成功后主面板

If Window("EC10.0").Window("Window").Exist(5) Then
	Reporter.ReportEvent micPass, "当前是否处于主面板","处于主面板！"	
Else
	Reporter.ReportEvent micFail, "当前是否处于主面板","不处于主面板！"
	Call ExitRun()
End If

'==================================
'perform based on subtype
'==================================
Select Case Lcase(strType)
	Case Lcase("主面板操作")
		Select Case Lcase(strSubType)
		'===============1.单击头像动作===================
			Case Lcase("单击头像")		
				Window("EC10.0").Window("Window").Click 26,37,micLeftBtn					
				If Window("资料设置").InsightObject("资料设置页").Exist(5)  Then
					Reporter.ReportEvent micPass,"资料页是否已经打开","是的资料页已经打开！"	
				else
				Window("EC10.0").Window("Window").Click 26,37,micLeftBtn
				  If Window("资料设置").InsightObject("资料设置页").Exist(5)  Then
					Reporter.ReportEvent micPass,"是否成功打开资料页","已成功打开资料页！"	
					else
					Reporter.ReportEvent micFail,"是否成功打开资料页","未成功打开资料页！重试打开一次。"
					Window("EC10.0").Window("Window").Click 26,37,micLeftBtn
					Window("资料设置").InsightObject("资料设置页").Exist(5)
					Reporter.ReportEvent micPass,"第二次点击是否成功打开资料页","已成功打开资料页！"
					End if
				End If
		'===============2.单击创建，进入创建群页面===================
			Case Lcase("进入创建群页")		
				Window("EC10.0").InsightObject("创建按钮").Click
				Window("EC10.0").InsightObject("创建-创建群").Click
				If Window("创建群").Window("Window").Exist(5)  Then
					Reporter.ReportEvent micPass,"是否成功打开创建群页","已成功打开创建群页！"	
				else
				Window("EC10.0").InsightObject("创建按钮").Click
				Window("EC10.0").InsightObject("创建-创建群").Click
				   If Window("创建群").Window("Window").Exist(5)  Then
					 Reporter.ReportEvent micPass,"是否成功打开创建群页","已成功打开创建群页！"	
				     else
				     Reporter.ReportEvent micFail,"是否成功打开创建群页","未成功打开创建群页！重试打开一次。"
				     Window("EC10.0").InsightObject("创建按钮").Click
				     Window("EC10.0").InsightObject("创建-创建群").Click
				     Window("创建群").Window("Window").Exist(5)
					 Reporter.ReportEvent micPass,"第二次点击是否成功打开创建群页","已成功打开创建群页！"
				     End if
				End If
		'===============3.单击创建，进入创建讨论组页面===================
			Case Lcase("进入创建讨论组页")	
				indow("EC10.0").InsightObject("创建按钮").Click
				Window("EC10.0").InsightObject("创建-创建讨论组").Click
				If Window("讨论组管理").InsightObject("创建讨论组页").Exist(5) Then
					Reporter.ReportEvent micPass,"是否成功打开创建讨论组页","已成功打开创建讨论组页！"	
				else
				Window("EC10.0").InsightObject("创建按钮").Click
				Window("EC10.0").InsightObject("创建-创建讨论组").Click
				If Window("讨论组管理").InsightObject("创建讨论组页").Exist(5) Then
					Reporter.ReportEvent micPass,"是否成功打开创建讨论组页","已成功打开创建讨论组页！"	
				     else
				     Reporter.ReportEvent micFail,"是否成功打开创建讨论组页","未成功打开创建讨论组页！重试打开一次。"
				     Window("EC10.0").InsightObject("创建按钮").Click
				     Window("EC10.0").InsightObject("创建-创建讨论组").Click
				     Window("讨论组管理").InsightObject("创建讨论组页").Exist(5)
					 Reporter.ReportEvent micPass,"第二次点击是否成功打开创建讨论组页","已成功打开创建讨论组页！"
				     End if
				End If
		'===============4.单击创建，进入添加好友页面===================
			Case Lcase("进入添加好友页")	
				window("EC10.0").InsightObject("创建按钮").Click
				Window("EC10.0").InsightObject("创建-添加好友").Click

				If Window("添加好友").InsightObject("添加好友页").Exist(5) Then
					Reporter.ReportEvent micPass,"是否成功打开添加好友页","已成功打开添加好友页！"	
				else
				Window("EC10.0").InsightObject("创建按钮").Click
				Window("EC10.0").InsightObject("创建-添加好友").Click
				     If Window("添加好友").InsightObject("添加好友页").Exist(5) Then
					Reporter.ReportEvent micPass,"是否成功打开添加好友页","已成功打开添加好友页！"	
				     else
				     Reporter.ReportEvent micFail,"是否成功打开添加好友页","未成功打开添加好友页！重试打开一次。"
				     Window("EC10.0").InsightObject("创建按钮").Click
				     Window("EC10.0").InsightObject("创建-添加好友").Click
				     Window("添加好友").InsightObject("添加好友页").Exist(5)
					 Reporter.ReportEvent micPass,"第二次点击是否成功打开添加好友页","已成功打开添加好友页！"
				     End if
				End If
		'===============5.单击会话列表，进入会话列表===================
			Case Lcase("进入会话列表")	
			If Window("EC10.0").InsightObject("会话列表激活状态").Exist(5) Then
					Reporter.ReportEvent micPass,"会话列表是否已经打开","已经打开！"	
				else
				Window("EC10.0").InsightObject("会话列表入口").Click
				Do until Window("EC10.0").InsightObject("会话列表激活状态").Exist(5)
					wait 1
				Loop
				End If
	  	'===============6.单击联系人列表，进入联系人列表===================
			Case Lcase("进入联系人列表")	
				If Window("EC10.0").InsightObject("联系人列表激活状态").Exist(5) Then
					Reporter.ReportEvent micPass,"联系人列表是否已经打开","已经打开！"	
				else
				Window("EC10.0").InsightObject("联系人列表入口").Click

				End If
	  	'===============7.单击我的客户，进入客户管理页===================
			Case Lcase("进入我的客户页")	
				If Window("EC10.0").InsightObject("我的客户激活状态").Exist(5) Then
					Reporter.ReportEvent micPass,"我的客户页是否已经打开","已经打开！"	
				else
				Window("EC10.0").InsightObject("我的客户入口").Click
				End If
		'===============8.单击销售助手，打开销售助手===================
			Case Lcase("进入销售助手")	
				If Window("销售助手").InsightObject("销售助手页检查点").Exist(5) Then
					Reporter.ReportEvent micPass,"销售助手页是否已经打开","已经打开！"	
				else
				Window("EC10.0").InsightObject("销售助手入口").Click

				End If	
		'===============9.单击电话，打开电话窗口===================
			Case Lcase("打开电话页面")	
				If Window("电话窗口").Window("Window").Exist(5) Then
					Reporter.ReportEvent micPass,"电话页面是否已经打开","已经打开！"	
				else
				Window("EC10.0").InsightObject("电话功能入口").Click
				End If		
		'===============10.单击短信，打开短信窗口===================
			Case Lcase("打开短信页面")	
				If Window("短信助手").Exist(5) Then
					Reporter.ReportEvent micPass,"短信页面是否已经打开","已经打开！"	
				else
				Window("EC10.0").InsightObject("短信助手入口").Click
				End If	
		'===============11.单击应用中心入口，打开应用中心页===================
			Case Lcase("打开应用中心页面")	
				If Window("EC10.0").InsightObject("应用中心激活状态").Exist(5) Then
					Reporter.ReportEvent micPass,"应用中心页面是否已经打开","已经打开！"	
				else
				Window("EC10.0").InsightObject("应用中心入口").Click

				End If	
		'===============12.单击收藏入口，打开收藏页===================
			Case Lcase("打开收藏页")	
				If Window("我的收藏").Window("Window").Exist(5) Then
					Reporter.ReportEvent micPass,"收藏页面是否已经打开","已经打开！"	
				else
				Window("EC10.0").InsightObject("收藏入口").Click
				End If	
		End Select
End Select