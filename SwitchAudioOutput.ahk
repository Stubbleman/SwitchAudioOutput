!a::
Devices := {}
IMMDeviceEnumerator := ComObjCreate("{BCDE0395-E52F-467C-8E3D-C4579291692E}", "{A95664D2-9614-4F35-A746-DE8DB63617E6}")

; IMMDeviceEnumerator::EnumAudioEndpoints
; eRender = 0, eCapture, eAll
; 0x1 = DEVICE_STATE_ACTIVE
DllCall(NumGet(NumGet(IMMDeviceEnumerator+0)+3*A_PtrSize), "UPtr", IMMDeviceEnumerator, "UInt", 0, "UInt", 0x1, "UPtrP", IMMDeviceCollection, "UInt")
ObjRelease(IMMDeviceEnumerator)

; IMMDeviceCollection::GetCount
DllCall(NumGet(NumGet(IMMDeviceCollection+0)+3*A_PtrSize), "UPtr", IMMDeviceCollection, "UIntP", Count, "UInt")
Loop % (Count)
{
    ; IMMDeviceCollection::Item
    DllCall(NumGet(NumGet(IMMDeviceCollection+0)+4*A_PtrSize), "UPtr", IMMDeviceCollection, "UInt", A_Index-1, "UPtrP", IMMDevice, "UInt")

    ; IMMDevice::GetId
    DllCall(NumGet(NumGet(IMMDevice+0)+5*A_PtrSize), "UPtr", IMMDevice, "UPtrP", pBuffer, "UInt")
    DeviceID := StrGet(pBuffer, "UTF-16"), DllCall("Ole32.dll\CoTaskMemFree", "UPtr", pBuffer)

    ; IMMDevice::OpenPropertyStore
    ; 0x0 = STGM_READ
    DllCall(NumGet(NumGet(IMMDevice+0)+4*A_PtrSize), "UPtr", IMMDevice, "UInt", 0x0, "UPtrP", IPropertyStore, "UInt")
    ObjRelease(IMMDevice)

    ; IPropertyStore::GetValue
    VarSetCapacity(PROPVARIANT, A_PtrSize == 4 ? 16 : 24)
    VarSetCapacity(PROPERTYKEY, 20)
    DllCall("Ole32.dll\CLSIDFromString", "Str", "{A45C254E-DF1C-4EFD-8020-67D146A850E0}", "UPtr", &PROPERTYKEY)
    NumPut(14, &PROPERTYKEY + 16, "UInt")
    DllCall(NumGet(NumGet(IPropertyStore+0)+5*A_PtrSize), "UPtr", IPropertyStore, "UPtr", &PROPERTYKEY, "UPtr", &PROPVARIANT, "UInt")
    DeviceName := StrGet(NumGet(&PROPVARIANT + 8), "UTF-16")    ; LPWSTR PROPVARIANT.pwszVal
    DllCall("Ole32.dll\CoTaskMemFree", "UPtr", NumGet(&PROPVARIANT + 8))    ; LPWSTR PROPVARIANT.pwszVal
    ObjRelease(IPropertyStore)

    ObjRawSet(Devices, DeviceName, DeviceID)
}
ObjRelease(IMMDeviceCollection)

; Initail device count number
DeviceCount := 1

Devices2 := {}
For DeviceName, DeviceID in Devices
    ObjRawSet(Devices2, A_Index, DeviceID)
    DeviceCount := DeviceCount + 1
    ; List .= "(" . A_Index . ") " . DeviceName . "`n", ObjRawSet(Devices2, A_Index, DeviceID)
; InputBox n, Audio Output, % List,,,,,,,, 1

Gui, New
Gui, Margin, 0, 0
Gui, -Caption
Gui, +AlwaysOnTop
Gui, Color, black,black
Gui, font, s12 w700 CWhite, Verdana

; Create the ListView with a device column.
; Raw number is given by device count.
; E0x200 = WS_EX_CLIENTEDGE ( Extended style)
Gui, Add, ListView, r%DeviceCount%  gAudioListView AltSubmit -E0x200 vnewobject BackgroundTrans -Border, Audio Device


For DeviceName, DeviceID in Devices
    LV_Add(, DeviceName)
    ; List .= "(" . A_Index . ") " . DeviceName . "`n", ObjRawSet(Devices2, A_Index, DeviceID)

Gui, Show

WinSet,Transparent,180,SwitchAudioOutput.ahk
WinGetPos, X, Y, W, H, A ; Retreive positoin and size of active window
WinSet, Region,0-0 w%W% h%H% R30-30,SwitchAudioOutput.ahk

; MsgBox, Size is %W% x %H%

KeyWait, Enter, D

Gui, Destroy

return

AudioListView:
Critical ,On　; Increase reliability, ensures that all "I" notifications are received.
if (A_GuiEvent = "I")
{
    ;  InStr(ErrorLevel, "S", true)
    ; LV_GetText(index, A_EventInfo)  ; Get the text from the row's first field.
    ; LV_GetText(index, ErrorLevel)  ; Get the text from the row's first field.
    index:= LV_GetNext(0, "F") ; Get selected column index
    ; LV_GetText(index, A_EventInfo , 1)
    ; MsgBox, , Title, %index% , 999
    ; MsgBox % ErrorLevel
    ; ToolTip % index

    ; Set audio output device with index.
    IPolicyConfig := ComObjCreate("{870af99c-171d-4f9e-af0d-e63df40c2bc9}", "{F8679F50-850A-41CF-9C72-430F290290C8}") ;00000102-0000-0000-C000-000000000046 00000000-0000-0000-C000-000000000046
    R := DllCall(NumGet(NumGet(IPolicyConfig+0)+13*A_PtrSize), "UPtr", IPolicyConfig, "Str", Devices2[index], "UInt", 0, "UInt")
    ObjRelease(IPolicyConfig)
}
return
