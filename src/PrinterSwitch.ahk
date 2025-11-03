; PrinterSwitch.ahk - 快速切换默认打印机工具 (AHK v2)
#Include Printer.ahk
#SingleInstance Force
Persistent

; 全局变量
global VERSION := 'v1.0.0 2025/11/03'
global ConfigFile := A_ScriptDir "\printers.ini"
global GuiWidth := 60
global GuiHeight := 60
global Printers := Array()
global GuiVisible := true  ; 跟踪GUI可见状态

; 创建配置文件（如果不存在）
CreateConfigFile()

; 从配置文件加载打印机列表
LoadPrintersFromConfig()

; 创建GUI窗口
CreatePrinterGui()

; 创建托盘菜单
CreateTrayMenu()

; 主循环
SetTimer UpdatePrinterStatus, 1000
return

; 创建GUI窗口
CreatePrinterGui() {
    global

    ; 创建主窗口
    GuiMain := Gui("+AlwaysOnTop -ToolWindow +LastFound +E0x20 -Caption")
    GuiMain.Title := "PrinterSwitch"
    GuiMain.BackColor := "0x36393F"
    GuiMain.MarginX := 4
    GuiMain.MarginY := 10

    ; 添加文本控件显示当前打印机简称
    currentPrinter := GetCurrentPrinterAlias()
    global CurrentPrinterText := GuiMain.Add("Text", "w" GuiWidth " h" GuiHeight " Left BackgroundTrans", currentPrinter)
    CurrentPrinterText.SetFont("cWhite s15 bold", "Arial")

    ; 设置窗口形状为圆角矩形
    WinSetRegion("0-0 w" GuiWidth " h" GuiHeight " R10-10", GuiMain.Hwnd)

    ; 显示窗口在屏幕右下角
    xPos := A_ScreenWidth - GuiWidth - 20
    yPos := A_ScreenHeight - GuiHeight - 20
    GuiMain.Show("NoActivate x" xPos " y" yPos)

    ; 绑定鼠标事件
    OnMessage(0x0201, WM_LBUTTONDOWN) ; 左键按下
    OnMessage(0x0204, WM_RBUTTONDOWN) ; 右键按下

    global PrinterGui := GuiMain
    ;UpdateIcon()
}

; 创建托盘菜单
CreateTrayMenu() {
    Tray := A_TrayMenu
    Tray.Delete() ; 删除默认菜单项
    tray.Add(VERSION, (*) => {})
    tray.Disable(VERSION)
    Tray.Add("隐藏/显示图标", ToggleGuiVisibility)
    Tray.Add("退出", (*) => ExitApp())
    ;Tray.Default := "隐藏/显示图标"
}

; 切换GUI可见性
ToggleGuiVisibility(*) {
    global
    if (GuiVisible) {
        PrinterGui.Hide()
        GuiVisible := false
    } else {
        PrinterGui.Show()
        GuiVisible := true
    }
}

; 处理左键按下事件（用于拖动窗口）
WM_LBUTTONDOWN(wParam, lParam, msg, hwnd) {
    global
    if (hwnd = PrinterGui.Hwnd) {
        PostMessage 0x00A1, 2, 0, , "ahk_id " hwnd
    }
}

; 处理右键按下事件（显示菜单）
WM_RBUTTONDOWN(wParam, lParam, msg, hwnd) {
    global
    if (hwnd = PrinterGui.Hwnd) {
        ShowPrinterMenu()
    }
}

; 显示打印机菜单
ShowPrinterMenu() {
    global

    ; 创建新菜单
    PrinterMenu := Menu()

    ; 添加每个打印机到菜单
    for _, prt in Printers {
        PrinterMenu.Add(prt.Alias, SetDefaultPrinter)
    }

    ; 显示菜单
    PrinterMenu.Show()
}

; 设置默认打印机
SetDefaultPrinter(ItemName, ItemPos, MenuName) {
    global

    ; 根据别名找到真实打印机名称
    prt := FindPrinterByAlias(ItemName)
    if (prt) {
        ; 设置为默认打印机
        DllCall("winspool.drv\SetDefaultPrinterW", "Str", prt.Name)

        ; 更新显示
        UpdateIcon()
    }
}

; 根据别名查找打印机对象
FindPrinterByAlias(alias) {
    global
    for _, prt in Printers {
        if (prt.Alias = alias) {
            return prt
        }
    }
    return ""
}

; 根据名称查找打印机对象
FindPrinterByName(name) {
    global
    for _, prt in Printers {
        if (prt.Name = name) {
            return prt
        }
    }
    return ""
}

; 更新GUI显示文本
UpdateIcon() {
    global
    alias := GetCurrentPrinterAlias()
    CurrentPrinterText.Text := alias
}

; 定时更新检查
UpdatePrinterStatus() {
    UpdateIcon()
}

; 获取当前默认打印机的别名
GetCurrentPrinterAlias() {
    ; 获取默认打印机名称
    size := 256
    VarSetStrCapacity(&printerName, size)
    success := DllCall("winspool.drv\GetDefaultPrinterW", "str", printerName, "UIntP", &size)

    if (success) {
        ; 获取到的是打印机的完整名称
        fullName := printerName
        ; msgbox fullName  ; 调试用，可以删除
        prt := FindPrinterByName(fullName)
        if (prt) {
            return prt.Alias
        }
        return SubStr(fullName, 1, 4) ; 如果未在配置中找到，返回前N个字符
    }
    return "未知"
}

; 从配置文件加载打印机列表
LoadPrintersFromConfig() {
    global ConfigFile, Printers
    Printers := Array() ; 清空现有列表

    ; 读取Printers section下的所有内容
    content := IniRead(ConfigFile, "Printers")
    if (content != "ERROR") {
        content := StrReplace(content, "`r", "")
        lines := StrSplit(content, "`n")

        ; 遍历每一行，创建Printer对象
        for _, line in lines {
            ; 过滤掉空行和注释行
            if (line != "" && !RegExMatch(line, "^\s*;")) {
                parts := StrSplit(line, "=")
                if (parts.Length >= 2) {
                    name := Trim(parts[1])
                    alias := Trim(parts[2])
                    if (name != "" && alias != "") {
                        Printers.Push(Printer(name, alias))
                    }
                }
            }
        }
    }
}

; 创建初始配置文件
CreateConfigFile() {
    global ConfigFile

    if (!FileExist(ConfigFile)) {
        ; 写入示例配置
        content .= "; 在这里添加打印机名称和对应简称的映射关系`n"
        content .= "; 格式：完整打印机名称=简称`n"
        content := "[Printers]`n"
        FileAppend(content, ConfigFile)
    }
}

; GUI退出事件
GuiMain_Close(*) {
    ExitApp
}
