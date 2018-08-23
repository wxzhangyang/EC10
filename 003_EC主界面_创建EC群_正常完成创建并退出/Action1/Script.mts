'**********************************************************************************************************************
'Test Name:				登陆页面_有效帐号密码直接登录_注销登录002
'ManualTestCaseName		V1.0_Story_06_T0003
'Purpose:					登录
'Premise Condition			

'Step
'1、点击账号输入框，输入账号
'2、点击密码输入框，输入密码
'3、点击登录按钮登录

'Expected Results

'Created by:    张阳                
'Creation date:2018.03.05
'-------------------------------------------------------------------------------------------
'**********************************************************************************************************************
'RunAction "Action1 [EC10_登录页面]", oneIteration, "输入登录账号"


RunAction "EC_10主界面 [EC10_主面板]", oneIteration,"单击创建进入创建群页"
RunAction "创建群 [EC10_创建群]", oneIteration,"输入群名"
RunAction "创建群 [EC10_创建群]", oneIteration,"选择只有管理员才能添加群成员"
RunAction "创建群 [EC10_创建群]", oneIteration,"点击下一步"
RunAction "创建群 [EC10_创建群]", oneIteration,"选择群成员完成创建"
RunAction "创建群 [EC10_创建群]", oneIteration,"检查创建成功"
RunAction "创建群 [EC10_创建群]", oneIteration,"解散群"
RunAction "创建群 [EC10_创建群]", oneIteration,"检查解散成功"
