﻿#Requires AutoHotkey v2.0+
#SingleInstance Force  ; The script will Reload if launched while already running
SetWorkingDir A_ScriptDir  ; Ensures a consistent starting directory
SendMode "Input"  ; Recommended for new scripts due to its superior speed and reliability

; Globals
DesktopCount := 2        ; Windows starts with 2 desktops at boot
CurrentDesktop := 1      ; Desktop count is 1-indexed (Microsoft numbers them this way)
LastOpenedDesktop := 1

; DLL
hVirtualDesktopAccessor := DllCall("LoadLibrary", "Str", A_ScriptDir . "\VirtualDesktopAccessor.dll", "Ptr")
global IsWindowOnDesktopNumberProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "IsWindowOnDesktopNumber", "Ptr")
global MoveWindowToDesktopNumberProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "MoveWindowToDesktopNumber", "Ptr")
global GoToDesktopNumberProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "GoToDesktopNumber", "Ptr")

; Main
SetKeyDelay, 75
mapDesktopsFromRegistry()
OutputDebug("[loading] desktops: " . DesktopCount . " current: " . CurrentDesktop)

#Include A_ScriptDir . "\user_config.ahk"
Return

; This function examines the registry to build an accurate list of the current virtual desktops and which one we're currently on.
mapDesktopsFromRegistry() {
    global CurrentDesktop, DesktopCount

    ; Get the current desktop UUID. Length should be 32 always, but there's no guarantee this couldn't change in a later Windows release so we check.
    IdLength := 32
    SessionId := getSessionId()
    if (SessionId) {
        RegRead, CurrentDesktopId, "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops", "CurrentVirtualDesktop"
        if (ErrorLevel) {
            RegRead, CurrentDesktopId, "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\SessionInfo\" . SessionId . "\VirtualDesktops", "CurrentVirtualDesktop"
        }

        if (CurrentDesktopId) {
            IdLength := StrLen(CurrentDesktopId)
        }
    }

    ; Get a list of the UUIDs for all virtual desktops on the system
    RegRead, DesktopList, "HKEY_CURRENT_USER", "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops", "VirtualDesktopIDs"
    if (DesktopList) {
        DesktopListLength := StrLen(DesktopList)
        ; Figure out how many virtual desktops there are
        DesktopCount := Floor(DesktopListLength / IdLength)
    } else {
        DesktopCount := 1
    }

    ; Parse the REG_DATA string that stores the array of UUID's for virtual desktops in the registry.
    i := 0
    while (CurrentDesktopId and i < DesktopCount) {
        StartPos := (i * IdLength) + 1
        DesktopIter := SubStr(DesktopList, StartPos, IdLength)
        OutputDebug("The iterator is pointing at " . DesktopIter . " and count is " . i)

        ; Break out if we find a match in the list. If we didn't find anything, keep the old guess and pray we're still correct :-D.
        if (DesktopIter = CurrentDesktopId) {
            CurrentDesktop := i + 1
            OutputDebug("Current desktop number is " . CurrentDesktop . " with an ID of " . DesktopIter)
            Break
        }
        i++
    }
}

; This function finds out the ID of the current session.
getSessionId() {
    ProcessId := DllCall("GetCurrentProcessId", "UInt")
    if (ErrorLevel) {
        OutputDebug("Error getting current process id: " . ErrorLevel)
        Return
    }
    OutputDebug("Current Process Id: " . ProcessId)

    DllCall("ProcessIdToSessionId", "UInt", ProcessId, "UInt*", SessionId)
    if (ErrorLevel) {
        OutputDebug("Error getting session id: " . ErrorLevel)
        Return
    }
    OutputDebug("Current Session Id: " . SessionId)
    Return SessionId
}

_switchDesktopToTarget(targetDesktop) {
    ; Globals variables should have been updated via updateGlobalVariables() prior to entering this function
    global CurrentDesktop, DesktopCount, LastOpenedDesktop

    ; Don't attempt to switch to an invalid desktop
    if (targetDesktop > DesktopCount or targetDesktop < 1 or targetDesktop == CurrentDesktop) {
        OutputDebug("[invalid] target: " . targetDesktop . " current: " . CurrentDesktop)
        Return
    }

    LastOpenedDesktop := CurrentDesktop

    ; Fixes the issue of active windows in intermediate desktops capturing the switch shortcut and therefore delaying or stopping the switching sequence. This also fixes the flashing window button after switching in the taskbar. More info: https://github.com/pmb6tz/windows-desktop-switcher/pull/19
    WinActivate("ahk_class Shell_TrayWnd")

    DllCall(GoToDesktopNumberProc, "Int", targetDesktop - 1)

    ; Makes the WinActivate fix less intrusive
    Sleep(50)
    focusTheForemostWindow(targetDesktop)
}

updateGlobalVariables() {
    ; Re-generate the list of desktops and where we fit in that. We do this because
    ; the user may have switched desktops via some other means than the script.
    mapDesktopsFromRegistry()
}

switchDesktopByNumber(targetDesktop) {
    global CurrentDesktop, DesktopCount
    updateGlobalVariables()
    _switchDesktopToTarget(targetDesktop)
}

switchDesktopToLastOpened() {
    global CurrentDesktop, DesktopCount, LastOpenedDesktop
    updateGlobalVariables()
    _switchDesktopToTarget(LastOpenedDesktop)
}

switchDesktopToRight() {
    global CurrentDesktop, DesktopCount
    updateGlobalVariables()
    _switchDesktopToTarget(CurrentDesktop == DesktopCount ? 1 : CurrentDesktop + 1)
}

switchDesktopToLeft() {
    global CurrentDesktop, DesktopCount
    updateGlobalVariables()
    _switchDesktopToTarget(CurrentDesktop == 1 ? DesktopCount : CurrentDesktop - 1)
}

focusTheForemostWindow(targetDesktop) {
    foremostWindowId := getForemostWindowIdOnDesktop(targetDesktop)
    if (isWindowNonMinimized(foremostWindowId)) {
        WinActivate("ahk_id " . foremostWindowId)
    }
}

isWindowNonMinimized(windowId) {
    WinGet, "MMX", "MinMax", "ahk_id " . windowId
    Return MMX != -1
}

getForemostWindowIdOnDesktop(n) {
    n := n - 1 ; Desktops start at 0, while in script it's 1

    ; winIDList contains a list of windows IDs ordered from the top to the bottom for each desktop.
    WinGet, winIDList, "list"
    Loop % winIDList {
        windowID := winIDList[A_Index]
        windowIsOnDesktop := DllCall(IsWindowOnDesktopNumberProc, "UInt", windowID, "UInt", n)
        ; Select the first (and foremost) window which is in the specified desktop.
        if (windowIsOnDesktop == 1) {
            Return windowID
        }
    }
}

MoveCurrentWindowToDesktop(desktopNumber) {
    WinGet, activeHwnd, "ID", "A"
    DllCall(MoveWindowToDesktopNumberProc, "UInt", activeHwnd, "UInt", desktopNumber - 1)
    switchDesktopByNumber(desktopNumber)
}

MoveCurrentWindowToRightDesktop() {
    global CurrentDesktop, DesktopCount
    updateGlobalVariables()
    WinGet, activeHwnd, "ID", "A"
    DllCall(MoveWindowToDesktopNumberProc, "UInt", activeHwnd, "UInt", (CurrentDesktop == DesktopCount ? 1 : CurrentDesktop + 1) - 1)
    _switchDesktopToTarget(CurrentDesktop == DesktopCount ? 1 : CurrentDesktop + 1)
}

MoveCurrentWindowToLeftDesktop() {
    global CurrentDesktop, DesktopCount
    updateGlobalVariables()
    WinGet, activeHwnd, "ID", "A"
    DllCall(MoveWindowToDesktopNumberProc, "UInt", activeHwnd, "UInt", (CurrentDesktop == 1 ? DesktopCount : CurrentDesktop - 1) - 1)
    _switchDesktopToTarget(CurrentDesktop == 1 ? DesktopCount : CurrentDesktop - 1)
}

createVirtualDesktop() {
    global CurrentDesktop, DesktopCount
    Send("#^d")
    DesktopCount++
    CurrentDesktop := DesktopCount
    OutputDebug("[create] desktops: " . DesktopCount . " current: " . CurrentDesktop)
}

deleteVirtualDesktop() {
    global CurrentDesktop, DesktopCount, LastOpenedDesktop
    Send("#^{F4}")
    if (LastOpenedDesktop >= CurrentDesktop) {
        LastOpenedDesktop--
    }
