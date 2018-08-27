'**********************************************************************************************************************
'Test Name:        销售助手 (ec主面板进入销售助手页后)
'Purpose:            
'Requirement:    
'Note:             
'Starting Page:    销售助手s
'Created by:        zhangyang                 
'Creation date:    14/08/2018
'
'Modification History: 
'Date:                Changed by:            Purpose:
'**********************************************************************************************************************
Option Explicit

'定义参数变量
Dim strType, strSubType, strNotes         '基础变量，不可修改
Dim strPlanName, strPlanNote,strCustomerName '参数化测试数据，需要自定义修改

'==================================================================
'依据业务脚本调用，确定执行步骤
Call SetActionTableRow(Parameter("Action"))
'==================================================================
'给变量赋值
strType = DataTable("Type", dtLocalSheet)
strSubType = DataTable("SubType", dtLocalSheet)
strNotes = DataTable("Notes", dtLocalSheet)

strPlanName = EvaluateInputParam(DataTable("计划标题", dtLocalSheet))
strPlanNote = EvaluateInputParam(DataTable("计划内容", dtLocalSheet))
strCustomerName = EvaluateInputParam(DataTable("客户姓名", dtLocalSheet))

Call CloseOptionalDialog(2)

'*****************************************************************
'脚本说明：检查当前是否在登录界面
'*****************************************************************
'If Window("EC10登录界面_操作项").InsightObject("登录按钮").Exist(5) Then
'    Reporter.ReportEvent micPass, "Verify Page - 登录界面","At 登录界面 Page"
'Else
'    Reporter.ReportEvent micFail, "Verify Page - 登录界面","Not at 登录界面 Page"
'    Call ExitRun()
'End If
'*****************************************************************
'脚本说明：登录界面可操作步骤
'*****************************************************************
Select Case Lcase(strType)
    Case Lcase("新建计划")
        Select Case Lcase(strSubType)
            '1.创建电话销售计划并验证提醒
            Case Lcase("创建电话销售计划并验证提醒")
               if Window("销售助手").InsightObject("新建计划").exist(5) then 
                  Window("销售助手").InsightObject("新建计划").Click
                 End if
                 Window("销售助手").InsightObject("电话计划创建入口").Click
                 if Window("新建计划").InsightObject("计划标题").exist(5) then 
                  Window("新建计划").InsightObject("计划标题").Type strPlanName & GetUniqueName()
                 End if
                '选择客户
                Window("新建计划").InsightObject("添加客户").Click
                wait 1

                Window("选择客户").InsightObject("搜索客户").type strCustomerName

                 Window("选择客户").DblClick 150,108,micLeftBtn 
                 Window("选择客户").InsightObject("确定").Click
                '填入电话计划内容
                Window("新建计划").InsightObject("电话计划内容").type strPlanNote
                 '设置时间
                 Window("新建计划").InsightObject("时间").Click
                 Window("新建计划").InsightObject("选择当前时间").Click
                 '完成创建
                 Window("新建计划").InsightObject("保存计划").Click
                If Window("Window").InsightObject("EC电话销售计划提醒").Exist(10) Then
                    Reporter.ReportEvent micPass, "是否正常弹出电话销售计划提醒?","正常弹出!"
                    Window("Window").InsightObject("关闭提醒").Click
                Else
                    Reporter.ReportEvent micFail, "是否正常弹出电话销售计划提醒?","没有弹出!"
                      call ReportErrorScreenShot("电话提醒弹窗","不正常")
                End If

            '2.新建短信计划
            Case Lcase("新建短信计划")
                  if Window("销售助手").InsightObject("新建计划").exist(5) then 
                  Window("销售助手").InsightObject("新建计划").Click
                 End if
                 Window("销售助手").InsightObject("短信计划创建入口").Click
                 if Window("新建计划").InsightObject("计划标题").exist(5) then 
                  Window("新建计划").InsightObject("计划标题").Type strPlanName & GetUniqueName()
                 End if
                 Window("新建计划").InsightObject("短信计划内容").type strPlanNote
                  '选择客户
                Window("新建计划").InsightObject("添加客户").Click
                wait 1

                Window("选择客户").InsightObject("搜索客户").type strCustomerName

                 Window("选择客户").DblClick 150,108,micLeftBtn 
                 Window("选择客户").InsightObject("确定").Click
                 '设置时间
                 Window("新建计划").InsightObject("时间").Click
                 Window("新建计划").InsightObject("选择当前时间").Click
                 '完成创建
                 Window("新建计划").InsightObject("保存计划").Click
                If Window("新建计划").InsightObject("保存成功").Exist(10) Then
                    Reporter.ReportEvent micPass, "是否正常保存短信计划?","正常保存!"
                Else
                    Reporter.ReportEvent micFail, "是否正常保存短信计划?","保存失败!"
                      call ReportErrorScreenShot("是否正常保存短信计划","保存失败")
                End If
            '3.新建QQ计划
              Case Lcase("新建QQ计划")
                if Window("销售助手").InsightObject("新建计划").exist(5) then 
                  Window("销售助手").InsightObject("新建计划").Click
                 End if
                 Window("销售助手").InsightObject("QQ计划创建入口").Click

                 if Window("新建计划").InsightObject("计划标题").exist(5) then 
                  Window("新建计划").InsightObject("计划标题").Type strPlanName & GetUniqueName()
                 End if
                 Window("新建计划").InsightObject("QQ计划内容").type strPlanNote
                  '选择客户
                Window("新建计划").InsightObject("添加客户").Click
                wait 1

                Window("选择客户").InsightObject("搜索客户").type strCustomerName

                 Window("选择客户").DblClick 150,108,micLeftBtn 
                 Window("选择客户").InsightObject("确定").Click
                 '设置时间
                 Window("新建计划").InsightObject("时间").Click
                 Window("新建计划").InsightObject("选择当前时间").Click
                 '完成创建
                 Window("新建计划").InsightObject("保存计划").Click
                If Window("SystemNoticePanelFrm").InsightObject("QQ计划执行校验").Exist(5) Then
                    Reporter.ReportEvent micPass, "是否正常执行QQ计划?","正常执行!"
                    Window("SystemNoticePanelFrm").InsightObject("关闭弹屏").Click
                Else
                    Reporter.ReportEvent micFail, "是否正常执行QQ计划?","执行失败!"
                      call ReportErrorScreenShot("是否正常保存短信计划","保存失败")
                End If
            '4.新建邮件计划
            Case Lcase("新建邮件计划")
                if Window("销售助手").InsightObject("新建计划").exist(5) then 
                  Window("销售助手").InsightObject("新建计划").Click
                 End if
                 Window("销售助手").InsightObject("邮件计划创建入口").Click
                 if Window("新建计划").InsightObject("计划标题").exist(5) then 
                  Window("新建计划").InsightObject("计划标题").Type strPlanName & GetUniqueName()
                 End if
                 Window("新建计划").InsightObject("邮件计划主题").type strPlanName & GetUniqueName()
                 Window("新建计划").InsightObject("邮件内容").type strPlanName & GetUniqueName()
                  '选择客户
                Window("新建计划").InsightObject("邮件计划添加客户").Click
                wait 1
                Window("选择客户").InsightObject("搜索客户").type strCustomerName

                 Window("选择客户").DblClick 150,108,micLeftBtn 
                 Window("选择客户").InsightObject("确定").Click
                 '设置时间
                 Window("新建计划").InsightObject("邮件计划时间").Click

                 Window("新建计划").InsightObject("选择当前时间").Click
                 '完成创建
                 Window("新建计划").InsightObject("保存计划").Click
                If Window("SystemNoticePanelFrm").InsightObject("邮件计划执行校验").Exist(5) Then
                    Reporter.ReportEvent micPass, "是否正常执行邮件计划?","正常执行!"
                    Window("SystemNoticePanelFrm").InsightObject("关闭弹屏").Click
                Else
                    Reporter.ReportEvent micFail, "是否正常执行QQ计划?","执行失败!"
                      call ReportErrorScreenShot("是否正常执行QQ计划","执行失败")
                End If
            '5.新建微信计划
            Case Lcase("新建微信计划")
                if Window("销售助手").InsightObject("新建计划").exist(5) then 
                  Window("销售助手").InsightObject("新建计划").Click
                 End if
                 Window("销售助手").InsightObject("微信计划创建入口").Click
                 if Window("新建计划").InsightObject("计划标题").exist(5) then 
                  Window("新建计划").InsightObject("计划标题").Type strPlanName & GetUniqueName()
                 End if
                 Window("新建计划").InsightObject("QQ计划内容").type strPlanNote
                  '选择客户
                Window("新建计划").InsightObject("添加客户").Click
                wait 1

                Window("选择客户").InsightObject("搜索客户").type strCustomerName

                 Window("选择客户").DblClick 150,108,micLeftBtn 
                 Window("选择客户").InsightObject("确定").Click
                 '设置时间
                 Window("新建计划").InsightObject("时间").Click
                 Window("新建计划").InsightObject("选择当前时间").Click
                 '完成创建
                 Window("新建计划").InsightObject("保存计划").Click
                If Window("新建计划").InsightObject("保存成功").Exist(5) Then
                    Reporter.ReportEvent micPass, "是否正常保存短信计划?","正常保存!"
                Else
                    Reporter.ReportEvent micFail, "是否正常保存短信计划?","保存失败!"
                      call ReportErrorScreenShot("是否正常保存短信计划","保存失败")
                End If
          
            '6.新建定时提醒
            Case Lcase("新建定时提醒")
                 if Window("销售助手").InsightObject("新建计划").exist(5) then 
                  Window("销售助手").InsightObject("新建计划").Click
                 End if
                 Window("销售助手").InsightObject("定时提醒创建入口").Click

                 if Window("新建计划").InsightObject("计划标题").exist(5) then 
                  Window("新建计划").InsightObject("计划标题").Type strPlanName & GetUniqueName()
                 End if
                 Window("新建计划").InsightObject("QQ计划内容").type strPlanNote
                  '选择客户
                Window("新建计划").InsightObject("添加客户").Click
                wait 1

                Window("选择客户").InsightObject("搜索客户").type strCustomerName

                 Window("选择客户").DblClick 150,108,micLeftBtn 
                 Window("选择客户").InsightObject("确定").Click
                 '设置时间
                 Window("新建计划").InsightObject("时间").Click
                 Window("新建计划").InsightObject("选择当前时间").Click
                 '完成创建
                 Window("新建计划").InsightObject("保存计划").Click
                If Window("新建计划").InsightObject("保存成功").Exist(10) Then
                    Reporter.ReportEvent micPass, "是否正常保存定时提醒?","正常保存!"
                    Window("SystemNoticePanelFrm").InsightObject("关闭弹屏").Click
                Else
                    Reporter.ReportEvent micFail, "是否正常保存定时提醒?","保存失败!"
                      call ReportErrorScreenShot("是否正常保存短信计划","保存失败")
                End If
           '7.关闭销售助手页
            Case Lcase("关闭销售助手页")
                 if Window("销售助手").InsightObject("关闭页面按钮").exist(2) Then
                    Window("销售助手").InsightObject("关闭页面按钮").Click
                    Else
                    reporter.ReportEvent micPass,"是否正常关闭销售助手页","已经关闭!"
                    End If
                   if Window("销售助手").InsightObject("关闭页面按钮").exist(2) Then

                    Reporter.ReportEvent micFail, "是否正常关闭销售助手页","未关闭,尝试再次关闭!"
                    window("销售助手").InsightObject("关闭页面按钮").Click
                    Call ExitRun()
                    else
                    reporter.ReportEvent micPass,"是否正常关闭销售助手页","已经关闭!"
                    Call ExitRun()
                   End If
               '8.销售助手页入口校验
              Case Lcase("销售助手页入口校验")
                 if Window("销售助手").InsightObject("销售助手窗口最大化未激活按钮").exist(2) Then 
                    Window("销售助手").InsightObject("销售助手窗口最大化未激活按钮").Click
                    Else
                    reporter.ReportEvent micFail,"最大化按钮是否存在?","最大化按钮不存在!"
                     Call ExitRun()
                    End If
                 if Window("销售助手").InsightObject("销售助手窗口最大化已激活按钮").exist(2) Then
                    Reporter.ReportEvent micPass, "销售助手窗口是否正常切换到最大化","正常切换到最大化!"
                    Window("销售助手").InsightObject("销售助手窗口最大化已激活按钮").Click
                    else
                    reporter.ReportEvent micFail,"销售助手窗口是否正常切换到最大化","没有正常切换到最大化!"
                    Call ExitRun()
                   End If
                 if Window("销售助手").InsightObject("销售助手窗口最大化未激活按钮").exist(2) Then 
                    Reporter.ReportEvent micPass, "销售助手窗口是否正常切换为初始窗口大小","正常切换!"
                    Else
                    reporter.ReportEvent micFail,"销售助手窗口是否正常切换为初始窗口大小?","没有正常切换!"
                     Call ExitRun()
                    End If
                 If Window("销售助手").InsightObject("销售助手窗口最小化按钮").Exist(2) Then
                 	Reporter.ReportEvent micPass, "销售助手页面最小化按钮是否存在","最小化按钮存在!"
                    else
                    reporter.ReportEvent micFail,"销售助手页面最小化按钮是否存在","最小化按钮不存在!"
                    Call ExitRun()
                    End if
                 If Window("销售助手").InsightObject("销售助手窗口设置入口").Exist(2) Then
                    Window("销售助手").InsightObject("销售助手窗口设置入口").Click
                    End if
                    If Window("销售助手").InsightObject("销售助手设置打开页").exist(2) Then
                 	Reporter.ReportEvent micPass, "销售助手设置是否打开","正常打开!"
                    else
                    reporter.ReportEvent micFail,"销售助手设置是否打开","未打开!"
                    Call ExitRun()
                    End If   
        End Select
End Select