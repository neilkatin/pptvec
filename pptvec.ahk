/*! pptvec -- adapted from example_2.ahk

   Take VEC footpedal events and allow them to the remapped

   May 2021, Neil Katin
*/
/*! TheGood
    AHKHID - An AHK implementation of the HID functions.
    AHKHID Example 2
    Last updated: August 22nd, 2010
    
    Registers HID devices and displays data coming upon WM_INPUT.
    This example shows how to use AHKHID_AddRegister(), AHKHID_Register(), AHKHID_GetInputInfo() and AHKHID_GetInputData().
    _______________________________________________________________
    1. Input the TLC (Usage Page and Usage) you'd like to register.
    2. Select any flags you want to associate with the TLC (see Docs for more info about each of them).
    3. Press Add to add the TLC to the array.
    3. Repeat 1, 2 and 3 for all the TLCs you'd like to register (the TLC array listview will get filled up).
    4. Press Call to register all the TLCs in the array.
    5. Any TLCs currently registered will show up in the Registered devices listview.
    6. Any data received will be displayed in the listbox.
    
    For example, if you'd like to register the keyboard and the mouse, put UsagePage 1 and check the flag RIDEV_PAGEONLY.
    Then press Add and then Call to register.
*/

#SingleInstance Force
#Include AHKHID/AHKHID.ahk
SetTitleMatchMode, 3
debugWindow := 0
debug := 0

; Gui window used for debug output
Gui +LastFound -Resize -MaximizeBox -MinimizeBox
Gui, Font, w700 s8, Courier New
Gui, Add, ListBox, x6 y6 w900 h320 vlbxInput hwndhlbxInput glbxInput_Event,

; look for command line arguments
if (A_Args.Length() > 0) {
    for n, param in A_Args
    {
        if (param == "--debugwindow") {
            debugWindow := 1
        } else if (param == "--debug") {
            debug := 1
        } else {
            MsgBox Parameter %n% is unknown: '%param%'.  Exiting...
            ExitApp
        }
    }
}



;Keep handle
GuiHandle := WinExist()

;Set up the constants
AHKHID_UseConstants()

;Intercept WM_INPUT
OnMessage(0x00FF, "InputMsg")

; magic to detect events from a VEC footpedal (Usagepage 12, Usage 3, get events even if not in foreground)
AHKHID_AddRegister(1)
AHKHID_AddRegister(12, 3, GuiHandle, RIDEV_INPUTSINK)
AHKHID_Register()

; assume no pedals down at the start
vecPedalDownLeft := 0
vecPedalDownMiddle := 0
vecPedalDownRight := 0

if (debugWindow) {
    Gui, Show
}
Return

Debug(string) {
    global lbxInput
    global hlbxInput
    global debugWindow
    global debug
    if (debugWindow) {
        GuiControl,, lbxInput, % "DEBUG: " . string
        SendMessage, 0x018B, 0, 0,, ahk_id %hlbxInput%                      ; LB_GETCOUNT
        SendMessage, 0x0186, ErrorLevel - 1, 0,, ahk_id %hlbxInput%         ; LB_SETCURSEL
    }
    if (debug) {
        FileAppend, % string "`n", **
    }
}

; called if close button in debug window is clicked
GuiClose:
ExitApp

; called if a row in the debug window is doubleclicked: clear the list
lbxInput_Event:
    If (A_GuiEvent = "DoubleClick") {
        GuiControl,, lbxInput,|
        iInputNum := 0
    }
Return


; called on foot pedal events
InputMsg(wParam, lParam) {
    Local r, h
    Critical    ;Or otherwise you could get ERROR_INVALID_HANDLE
    
    h := AHKHID_GetInputInfo(lParam, II_DEVHANDLE)
    r := AHKHID_GetInputData(lParam, uData)

    byte1 := NumGet(uData, 1, "UChar")                                   ; origin zero

    DetectPedalDown(vecPedalDownLeft, "left", byte1 & 0x1)
    DetectPedalDown(vecPedalDownMiddle, "middle", byte1 & 0x2)
    DetectPedalDown(vecPedalDownRight, "right", byte1 & 0x4)
}


; called on each event for each pedal (left, middle, right).
; track state changes; record current state of each pedal
; call NewDownEvent() if a pedal press is detected
DetectPedalDown(ByRef state, name, value) {

    if (value != 0) {
        ; pedal is now down
        if (state != 0) {
            ; pedal was already down: ignore
        } else {
            ; this is a new down event
            ;Debug("new pedal down " . name . " old state '" . state . "' value '" . value . "'")
            NewDownEvent(name)
        }
    } else {
        ; pedal is not down: reset the state unconditionally
        if (state != 0) {
            ;Debug("resetting pedal " . name . " old state '" . state . "' value '" . value . "'")
        }
    }
    state := value
}

; called when a new pedal press is detected.  Type is left, middle, or right
NewDownEvent(type) {

    Debug("NewDownEvent called: type " . type)
    windowClass0 := "ahk_class screenClass"
    windowClass1 := "ahk_class PPTFrameClass"
    controlName1 := "mdiClass1"

    keyToSend := ""
    if (type = "left" ) {
        keyToSend := "{PgUp}"
    } else if (type = "middle") {
        keyToSend := "{PgDn}"
    }

    ; the problem: pgup/down events don't work if ppt is in full screen mode.  Sending com events does work though (but not if ppt
    ; is in window mode, so we need two differnt code paths

    ; controlName is blank if we are in full screen mode; try to use com commands instead
    if (WinExist(windowClass0)) {
        ppt := ComObjActive("PowerPoint.Application")
        if (ppt) {
            try {
                ; thanks to reddit user daonlyfreez for com advance suggestion
                ; https://www.reddit.com/r/AutoHotkey/comments/bqa0fc/forwardback_on_a_powerpoint_2016_presentation/eo4bgvg
                if (type = "left") {
                    Debug("sending com prevous")
                    ppt.ActivePresentation.SlideShowWindow.View.Previous
                } else if (type = "middle") {
                    Debug("sending com next")
                    ppt.ActivePresentation.SlideShowWindow.View.Next
                }
                
                ; don't send page up/down -- we already advanced the page
                keyToSend := ""
            } catch e {
                Debug("NewDownEvent: got an exception when using com: '" . e . "'")
            }
        }
    }

    ; send Page up/down when not in full screen mode
    if (keyToSend != "") {
        if (WinExist(windowClass1)) {
            Debug("sending key '" keyToSend "' to window '" windowClass1 "' control '" controlName1 "'")
            try {
                ControlSend , %controlName1%, %keyToSend%, %windowClass1%
            } catch e {
                Debug("NewDownEvent: got an exception when sending key: '" . e . "'")
            }
        }
    }

    return 1
}

