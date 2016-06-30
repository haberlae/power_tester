import os
import autoit
import time
from datetime import datetime
from easygui import msgbox


def make_pack_folder(p_number, w_speed, w_style):
    root = "C:\\Users\\Eric\\Desktop\\15sec_Capture\\LP Power Data\\"
    target = root + "LP-{0}".format(p_number)
    if not os.path.exists(target):
        print(" ")
        print("(made new pack folder at: ", target)
        os.mkdir(target)
    return make_speed_folder(target, w_speed, w_style)


def make_speed_folder(s_target, w_speed, w_style):
    speed_target = os.path.join(s_target, "{0}MPH-{1} Walk".format(w_speed, w_style))
    if not os.path.exists(speed_target):
        os.mkdir(speed_target)
    else:
        return


def main():
    for i in range(0, 2):
        print(".")
    print("Please enter the info below for the pack to be tested: ")
    print(" ")
    pack_number = input("Pack Number: ")
    print(" ")
    walk_speed = input("Test Speed: ")
    print(" ")
    walk_style = input("Walk Style: ")

    make_pack_folder(pack_number, walk_speed, walk_style)

    print('.....')
    print("..........Launching.....")

    autoit.auto_it_set_option('MouseCoordMode', 0)
    autoit.auto_it_set_option('SendKeyDelay', 10)
    autoit.run("C:\\Program Files (x86)\\Pico Technology\\PicoScope6\\PicoScope.exe")
    autoit.win_wait('PicoScope 6')
    autoit.win_activate('PicoScope 6')
    autoit.send('{ALT}{f}{o}')
    # autoit.control_click('Open', "[Class:Edit;INSTANCE:1]", "")
    time.sleep(2)
    autoit.send('C:\\Users\\Eric\\Desktop\\15sec_Capture\\electrical_power.psdata', 1)
    autoit.send('{ENTER}')
    # autoit.control_click('Open', "[Class:Button;INSTANCE:1]")
    autoit.win_wait('PicoScope 6 - [electrical_power.psdata]')
    autoit.win_activate('PicoScope 6 - [electrical_power.psdata]')
    time.sleep(0.5)
    autoit.send('{ALT}{TAB}{TAB}{TAB}{TAB}{ENTER}{END}{UP}{ENTER}')
    autoit.win_wait('Macro Recorder')
    autoit.control_click('Macro Recorder', "[Name:_buttonImport]")
    time.sleep(0.5)
    # autoit.control_click('Open', "[Class:Edit;INSTANCE:1]")
    # time.sleep(0.5)
    autoit.send('C:\\Users\\Eric\\Desktop\\15sec_Capture\\15sec_pico_record_macro.psmacro')
    autoit.send('{ENTER}')
    #autoit.control_click('Open', "[Class:Button;INSTANCE:1]")
    msgbox("Ready to go?")
    autoit.win_activate('Macro Recorder')
    autoit.win_wait('Macro Recorder')
    autoit.control_click('Macro Recorder', "[Name:_buttonExecute]")
    time.sleep(8)
    autoit.win_activate('Macro Recorder')
    autoit.send('{ESCAPE}')

    msgbox("Stop Runnin'-- Enter power numbers into spreadsheet. Then return to this box and hit OK to continue.")
    autoit.win_activate('PicoScope 6')
    autoit.send('{ALT}{f}{a}')
    time.sleep(.5)
    i = datetime.now()
    autoit.send(os.path.join("C:\\Users\\Eric\\Desktop\\15sec_Capture\\LP Power Data\\",
                             "LP-{0}".format(pack_number), "{0}MPH-{1} Walk".format(walk_speed, walk_style),
                             "{0}MPH-{1} Walk_".format(walk_speed, walk_style) + i.strftime('%Y-%m-%d %Hh%Mm%Ss')+".psdata"))
    time.sleep(1)
    autoit.send('{ENTER}')
    time.sleep(5)
    autoit.process_close('PicoScope.exe')

    print("........................... Testing Complete.")


if __name__ == '__main__':
    main()
