'本程序由@Z-L-Alfred编写，只供用于非商业用途，侵权必究！作者bilibili主页：https://space.bilibili.com/3493113890867592/
c=InputBox ("请输入要开平方的数。"&Chr(13)&"记得输入正数哦~"&Chr(13)&"提示： 0 的算数平方根是 0 ~~","迭代法开平方 - 4 步骤之 1 - 设置开平方数","1")
If IsNumeric(c)=True Then
If c>0 Then
d=InputBox ("请输入要迭代的次数。"&Chr(13)&"提示：次数越多，结果越精准~","迭代法开平方 - 4 步骤之 2 - 设置迭代次数","4")
If IsNumeric(d)=True Then
If d>=1 Then
a=1
For i=1 To d
b=c/a
a=(a+b)/2
MsgBox "第 "&i&" 次迭代的结果约是 "&a&" 。",0,"迭代法开平方 - 4 步骤之 3 - 迭代"
Next
MsgBox "经过 "&d&" 次迭代得到， "&c&" 的算术平方根约是 "&a&" 。",0,"迭代法开平方 - 4 步骤之 4 - 输出计算结果"
Else
MsgBox "您输入了一个不正确的迭代次数。"&Chr(13)+Chr(13)&"按“确定”关闭程序。",16,"迭代法开平方 - 错误"
End If
Else
MsgBox "您输入了一个不正确的迭代次数。"&Chr(13)+Chr(13)&"按“确定”关闭程序。",16,"迭代法开平方 - 错误"
End If
ElseIf c=0 Then
MsgBox "您仿佛没有正确输入要开平方的数。"&Chr(13)+Chr(13)&"按“确定”关闭程序。",16,"迭代法开平方 - 错误"
ElseIf c<0 Then
MsgBox "您输入了一个负数。"&Chr(13)&"在实数集内，负数是没有平方根的。"&Chr(13)+Chr(13)&"按“确定”关闭程序。",16,"迭代法开平方 - 错误"
End If
Else
MsgBox "您没有正确输入要开平方的数。"&Chr(13)+Chr(13)&"按“确定”关闭程序。",16,"迭代法开平方 - 错误"
End If