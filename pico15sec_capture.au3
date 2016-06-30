AutoItSetOption('MouseCoordMode', 0)
AutoItSetOption('SendKeyDelay', 10)
Run("C:\Program Files (x86)\Pico Technology\PicoScope6\PicoScope.exe")
WinWait('PicoScope 6')
WinActivate('PicoScope 6')
Send('{ALT}{f}{o}')
ControlClick('Open', '', "[Class:Edit;INSTANCE:1]")
Sleep(1000)
Send('C:\Users\Eric\Desktop\15sec_Capture\electrical_power.psdata', 1)
ControlClick('Open', '', "[Class:Button;INSTANCE:1]")
WinWait('PicoScope 6 - [electrical_power.psdata]')
WinActivate('PicoScope 6 - [electrical_power.psdata]')
Sleep(500)
Send('{ALT}{TAB}{TAB}{TAB}{TAB}{ENTER}{END}{UP}{ENTER}')
WinWait('Macro Recorder')
ControlClick('Macro Recorder', '',"[Name:_buttonImport]")
Sleep(300)
ControlClick('Open', '', "[Class:Edit;INSTANCE:1]")
Sleep(500)
Send('C:\Users\Eric\Desktop\15sec_Capture\15sec_pico_record_macro.psmacro')
ControlClick('Open', '', "[Class:Button;INSTANCE:1]")
WinWait('Macro Recorder')
ControlClick('Macro Recorder', '', "[Name:_buttonExecute]")
Sleep(17000)
WinActivate('Macro Recorder')
Send('{ESCAPE}')
MsgBox(48, "SCRIPT PAUSE", "Check your power numbers and enter them into spreadsheet. Then return to this box and press OK to continue.")
WinActivate('PicoScope 6')
Send('{ALT}{f}{a}')







