/*! pptvec -- adapted from example_2.ahk

   Take VEC footpedal events and allow them to the remapped
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

#Include ../AHKHID.ahk
SetTitleMatchMode, 3

;Check if the OS is Windows Vista or higher
bVista := (DllCall("GetVersion") & 0xFF >= 6)

;Create GUI
Gui +LastFound -Resize -MaximizeBox -MinimizeBox
Gui, Font, w700 s8, Courier New
Gui, Add, ListBox, x6 y6 w650 h320 vlbxInput hwndhlbxInput glbxInput_Event,

;Keep handle
GuiHandle := WinExist()

;Set up the constants
AHKHID_UseConstants()

;Intercept WM_INPUT
OnMessage(0x00FF, "InputMsg")

AHKHID_AddRegister(1)
AHKHID_AddRegister(12, 3, GuiHandle, RIDEV_INPUTSINK)
AHKHID_Register()

vecPedalDownLeft := 0
vecPedalDownMiddle := 0
vecPedalDownRight := 0

debugWindow := 1

;Show GUI
Gui, Show
Return

Debug(string) {
    global lbxInput
    global debugWindow
    if (debugWindow) {
        GuiControl,, lbxInput, % "DEBUG: " . string
        SendMessage, 0x018B, 0, 0,, ahk_id %hlbxInput%                      ; LB_GETCOUNT
        SendMessage, 0x0186, ErrorLevel - 1, 0,, ahk_id %hlbxInput%         ; LB_SETCURSEL
    }
}

GuiClose:
ExitApp

;Clear on doubleclick
lbxInput_Event:
    If (A_GuiEvent = "DoubleClick") {
        GuiControl,, lbxInput,|
        iInputNum := 0
    }
Return


InputMsg(wParam, lParam) {
    Local r, h
    Critical    ;Or otherwise you could get ERROR_INVALID_HANDLE
    
    h := AHKHID_GetInputInfo(lParam, II_DEVHANDLE)
    r := AHKHID_GetInputData(lParam, uData)
    ;;;GuiControl,, lbxInput, % ""
        ;;;. "Vendor ID: "   Format("0x{1:04x}", AHKHID_GetDevInfo(h, DI_HID_VENDORID,     True))
        ;;;. " Product ID: "  Format("0x{1:04x}", AHKHID_GetDevInfo(h, DI_HID_PRODUCTID,    True))
        ;;;. " UsPg/Us: " AHKHID_GetDevInfo(h, DI_HID_USAGEPAGE, True) . "/" . AHKHID_GetDevInfo(h, DI_HID_USAGE, True)
        ;;;. " Data: " Bin2Hex(&uData, r)
        ;;;. " Raw: bytes " . r . " content " . format("{1:x} {2:x} {3:x}", NumGet(uData, 0, "UChar"), NumGet(uData, 1, "UChar"), NumGet(uData, 2, "UChar"))

    byte1 := NumGet(uData, 1, "UChar")                                   ; origin zero

    DetectPedalDown(vecPedalDownLeft, "left", byte1 & 0x1, lbxInput)
    DetectPedalDown(vecPedalDownMiddle, "middle", byte1 & 0x2, lbxInput)
    DetectPedalDown(vecPedalDownRight, "right", byte1 & 0x4, lbxInput)

    ;GuiControl,, lbxInput, % "" "extra line"
    ;GuiControl,, lbxInput, % "" " "                                     ; blank line
    SendMessage, 0x018B, 0, 0,, ahk_id %hlbxInput%                      ; LB_GETCOUNT
    SendMessage, 0x0186, ErrorLevel - 1, 0,, ahk_id %hlbxInput%         ; LB_SETCURSEL
}


DetectPedalDown(ByRef state, name, value, ByRef labelVar) {

    ;GuiControl,, labelVar, % "" . "DetectPedalDown called.  Pedal " . name . " State " . format("{1:x}", state) . " Value " . format("{1:x}", value)

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

NewDownEvent(type) {

    Debug("NewDownEvent called: type " . type)
    windowClass := "PPTFrameClass"
    controlName := "mdiClass1"

    ; just for testing
    ;windowClass := "Notepad++"
    ;controlName := "Scintilla1"

    pptHandle := WinExist("ahk_class " windowClass)

    if (pptHandle) {
        ;Debug("NewDownEvent: ppt window handle found: '" . pptHandle . "'")

        keyToSend := ""
        if (type == "left" ) {
            keyToSend := "{PgUp}"
        } else if (type == "middle") {
            keyToSend := "{PgDn}"
        }

        if (keyToSend != "") {
            Debug("sending key " keyToSend " to window class " windowClass " control " controlName)
            try {
                ControlSend , %controlName%, %keyToSend%, ahk_class %windowClass%
            } catch e {
                Debug("NewDownEvent: got an exception when sending key: '" . e . "'")
            }
        }
    }
}

