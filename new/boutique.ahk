#Persistent  ; 使脚本保持运行
SetTitleMatchMode, 2  ; 设置标题匹配模式以便匹配部分标题

    ; 切换到 Excel 窗口（根据窗口标题适当调整）
    IfWinExist, info.xlsx - Excel
    {
        WinActivate
    }

; 获取 Excel 对象
xl := ComObjActive("Excel.Application")

; 获取或创建工作簿
wb := xl.ActiveWorkbook ; 获取当前激活的工作簿

; 获取或创建工作表
ws := wb.ActiveSheet ; 获取当前激活的工作表

; 选中特定单元格
Cell := ws.Range("A2") ; 替换为你要选择的单元格
Cell.Select()

; 等待一段时间，你可以根据需要调整等待时间
Sleep, 900


    ; 模拟 Ctrl+C 复制 名称
    Send, ^c
 Sleep, 300
    ; 切换到目标应用程序窗口（根据窗口标题适当调整）
    IfWinExist, Google Chrome
    {
        WinActivate
    }
    ; 等待一段时间以确保粘贴操作
    Sleep, 500
    ; 模拟 Ctrl+V 粘贴
    Send, ^v
 Sleep, 300
	
	; 切换到 Excel 窗口（根据窗口标题适当调整）
    IfWinExist, info.xlsx - Excel
    {
        WinActivate
    }
	Send, {Tab} 
    ; 模拟 Ctrl+C 复制 邮箱
    Send, ^c
 Sleep, 300

; 切换到目标应用程序窗口（根据窗口标题适当调整）
    IfWinExist, Google Chrome
    {
        WinActivate
    }
Send, {Tab} 

Send, {Tab}
    ; 等待一段时间以确保粘贴操作
    Sleep, 500
    ; 模拟 Ctrl+V 粘贴
    Send, ^v
    Sleep, 300

PerformCopyPaste() {

; 切换到 Excel 窗口（根据窗口标题适当调整）
    IfWinExist, info.xlsx - Excel
    {
        WinActivate
    }
	Send, {Tab} 
    ; 模拟 Ctrl+C 复制 电话
    Send, ^c
 Sleep, 300


; 切换到目标应用程序窗口（根据窗口标题适当调整）
    IfWinExist, Google Chrome
    {
        WinActivate
    }
Send, {Tab} 

    ; 等待一段时间以确保粘贴操作
    Sleep, 500
    ; 模拟 Ctrl+V 粘贴
    Send, ^v
  Sleep, 300

}

; 调用函数 5 次
Loop, 5
{
    PerformCopyPaste()
}

Send, {Tab} 
Send, {Tab} 

Send, {Enter}
 Sleep, 500

; 倒退至第一格子
Backup() {
	Send, +{Tab}
}

Loop,10

{
    Backup()
}

; 调用函数 7 次
Loop, 7
{
    PerformCopyPaste()
}

    return
Exit
