#Pyhton 3.x
# -*- coding: UTF-8 -*-

import time 
import traceback
import re
import os,sys


import win32con
import win32api
import win32clipboard

from tkinter import *
from tkinter import messagebox, tix as Tix
import tkinter.font as tf

import threading
from pynput import keyboard
from pynput.mouse import Listener as mouseListener

WindX  = {}
WindX['self_folder'] = re.sub(r'\\','/',os.path.abspath(os.path.dirname(__file__)))
print("\nroot:",WindX['self_folder'])  
sys.path.append(WindX['self_folder'])  
os.chdir(WindX['self_folder'])
WindX['pcName'] = os.environ['COMPUTERNAME']
print("getcwd:",os.getcwd() + "\nDevice Name:",WindX['pcName'])

WindX['main'] = None
WindX['Label'] = None
WindX['Ctrl_Pressed'] = 0
WindX['Ctrl_C'] = 0
WindX['check'] = 0
WindX['stop_cheking'] = 0
WindX['TopLevel'] = None

def CheckClipboard():
    if WindX['stop_cheking']:
        return

    WindX['check'] +=1
    WindX['Label'].configure(text='---------------- Checking: ' + str(WindX['check']) + " ----------------")
    WindX['main'].update()
    gs = re.split(r'x|\+', WindX['main'].geometry()) #506x152+-1418+224
    WindX['main'].geometry('+'+ str(gs[2]) +'+' + str(gs[3]))

    time.sleep(0.5)
    text = ''
    try:
        win32clipboard.OpenClipboard()
        text = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
    except:
        pass

    matchQTY = 0
    matchedTypes = {}

    if len(text):
        #print("\tClipboard=", text)
        FindWords = []
        FindWords = re.split(r'\s*\,\s*', re.sub(r'^\s+|\s+$', '', WindX['FindStringVar'].get()))
        print("\tFind Words=", FindWords)

        if len(FindWords):            
            row = 0
            ft = tf.Font(family='微软雅黑',size=10) ###有很多参数
            a = '1.0'
            WindX['Text'].tag_add('tag_match', a)
            WindX['Text'].tag_config('tag_match',foreground ='blue',background='pink',font = ft)

            b = '1.0'
            WindX['Text'].tag_add('tag_no_match', b)
            WindX['Text'].tag_config('tag_no_match',foreground ='black',font = ft)

            c = '1.0'
            WindX['Text'].tag_add('tag_match_word', c)
            WindX['Text'].tag_config('tag_match_word',foreground ='red',background='yellow',font = ft)

            d = '1.0'
            WindX['Text'].tag_add('tag_x', c)
            WindX['Text'].tag_config('tag_x',foreground ='black',background='white',font = ft)

            for t in re.split(r'\n', text):
                row +=1
                matches = []
                match_points = {}
                #print(t)
                tcopy = t
                for s in FindWords: 
                    #['Category','Group','Business\s+Sector','Audit\s+Finding\s+Type', 'Audit\s+Type', 'Severity', 'Division', 'CAR\s+Type', 'Concern\s+Type', 'Type', 'Checklist']:
                    matchObj = re.match(r'.*({})'.format(s), tcopy, re.I)
                    if matchObj:
                        #print(matchObj.groups())
                        matches.append(matchObj.groups()[0])
                        matchedTypes[matchObj.groups()[0]] = 1

                        st = 0
                        while st > -1:
                            index = tcopy.find(matchObj.groups()[0], st)
                            if index > -1:                            
                                match_points[index + 1] = [matchObj.groups()[0], index, index + len(matchObj.groups()[0]) - 1]
                                tcopy = tcopy.replace(matchObj.groups()[0],' '*len(matchObj.groups()[0]),1)
                                #print(tcopy)
                                st = index + len(matchObj.groups()[0])
                            else:
                                st = -1

                if len(matches):
                    matchQTY +=1 #len(matches)
                    #print(match_points) 
                    st = 0
                    for s in sorted(match_points.keys()):
                        print(s, match_points[s])

                        dt_b = ''
                        dt_a = ''
                        if match_points[s][1] > st:
                            dt_b = t[st : match_points[s][1]]
                            dt_a = t[match_points[s][1]: match_points[s][2]+1]
                        else:
                            dt_a = t[match_points[s][1]: match_points[s][2]+1]
                        
                        
                        #print(dt_b, '--', dt_a)

                        if dt_b:
                            b = str(row) + '.' + str(st)
                            WindX['Text'].insert(b, dt_b, 'tag_no_match')

                        if dt_a:
                            a = str(row) + '.' + str(match_points[s][1])
                            WindX['Text'].insert(a, dt_a, 'tag_match')

                        st = match_points[s][2] + 1
                        
                    if st < len(t) -1:
                        dt_b = t[st : len(t)]
                        b = str(row) + '.' + str(st)
                        WindX['Text'].insert(b, dt_b, 'tag_no_match')

                    #a = str(row) + '.0'                
                    #WindX['Text'].insert(a, t, 'tag_match')

                    c = str(row) + '.' + str(len(t) + 5)
                    ms= "   ::["+str(matchQTY)+"]:: Found ["+", ".join(matches)+"]  "
                    WindX['Text'].insert(c, ms, 'tag_match_word')
                    
                    d = str(row) + '.' + str(len(t + ms) + 5)
                    WindX['Text'].insert(d, "\n", 'tag_x')
                else:
                    b = str(row) + '.0'
                    WindX['Text'].insert(b, t + "\n", 'tag_no_match')

    if matchQTY:
        matches = sorted(matchedTypes.keys())
        msg = 'Found in ' + str(matchQTY) + " lines -- checked done! Found words: [" + ", ".join(matches) + "]"
        WindX['Status'].configure(text=msg,bg='yellow', fg='red')
        Message(msg, 'yellow', 'red')

    else:
        WindX['Status'].configure(text='No found -- checked done!',bg='#EFEFEF', fg='green')
        Message('No found -- checked done!', '#EFEFEF', 'green')


    WindX['main'].update()     

def Message(msg, bgColor, fgColor):
    WindX['TopLevel'] = Toplevel()
    WindX['TopLevel'].wm_attributes('-topmost',1)        
    pos = win32api.GetCursorPos()
    WindX['TopLevel'].geometry('+'+ str(pos[0]) +'+' + str(pos[1] - 20))
    WindX['TopLevel'].overrideredirect(1)

    font_type = None 
    try:
        font_type = tf.Font(family="Lucida Grande", size=15)
    except:
        pass
    label = Label(WindX['TopLevel'], text=msg, justify=LEFT, relief=FLAT,pady=3,padx=3, anchor='w', bg=bgColor, fg=fgColor, font=font_type)
    label.pack(side=TOP, fill=X)

def WindExit():           
    WindX['main'].destroy()    
    os._exit(0)
    #sys.exit(0)  # This will cause the window error: Python has stopped working ...

def main():
    WindX['main'] = Tix.Tk()
    WindX['main'].title("Find Words in Clipboard")
    WindX['main'].geometry('+0+30')
    #WindX['main'].wm_attributes('-topmost',1)
    WindX['main'].protocol("WM_DELETE_WINDOW", WindExit)

    WindX['Label'] = Label(WindX['main'], text='---------------- Check: 0 ----------------', justify=CENTER, relief=FLAT,pady=3,padx=3)
    WindX['Label'].pack(side=TOP, fill=X)

    frame1 = Frame(WindX['main'])
    frame1.pack(side=TOP, fill=X, pady=3,padx=3)
    Label(frame1, text='Find: ', justify=CENTER, relief=FLAT, pady=3, padx=3).pack(side=LEFT) 

    WindX['FindStringVar'] = StringVar()
    e=Entry(frame1, justify=LEFT, relief=FLAT, textvariable= WindX['FindStringVar'], fg='blue')
    e.pack(side=LEFT, fill=BOTH, pady=3,padx=0, expand=True)
    e.insert(0,'Category, Group, Business\s+Sector, Audit\s+Finding\s+Type, Audit\s+Type, Severity, Division, CAR\s+Type, Concern\s+Type, Type, Checklist')
    #WindX['Find_Entry'] = e

    framex = Frame(WindX['main'])
    #s1 = Scrollbar(framex, orient= HORIZONTAL)
    s2 = Scrollbar(framex, orient= VERTICAL)
    WindX['Text'] = Text(framex, padx=5, pady=5, yscrollcommand=s2.set, wrap=WORD)  #wrap=WORD  NONE CHAR  xscrollcommand=s1.set, 
    s2.config(command=WindX['Text'].yview)     
    #s1.config(command=WindX['Text'].xview)

    #s1.pack(side=BOTTOM,fill=X)
    s2.pack(side=RIGHT,fill=Y) 
    WindX['Text'].pack(side=LEFT, fill=BOTH, expand=True)
    framex.pack(side=TOP,fill=BOTH, expand=True)

    WindX['Text'].insert('1.0', "\n\t\tPress [Esc] key to pause, press again to continue.")

    WindX['Status'] = Label(WindX['main'], text='', justify=LEFT, relief=FLAT,pady=3,padx=3, anchor='w', fg='green')
    WindX['Status'].pack(side=TOP, fill=X)

    mainloop()

def on_press(key):
    try:
        if isinstance(key, keyboard.Key):
            #print('Key:', key.name, ",", key.value.vk)
            if key.name == 'ctrl_l':
                WindX['Ctrl_Pressed'] = 1 
                #print('special key {0} pressed'.format(key))
            elif key.name == 'esc':
                if WindX['stop_cheking']:
                    WindX['stop_cheking'] = 0
                    WindX['Label'].configure(text='---------------- Ready to check ----------------')
                    WindX['main'].update()
                else:
                    WindX['stop_cheking'] = 1
                    WindX['Label'].configure(text='---------------- locked now, hit [ESC] key to unlock it! ----------------')                    
                    WindX['main'].update()

        elif isinstance(key, keyboard.KeyCode):
            #print('KeyCode:', key.char,",", key.vk)
            #print(key.vk == 67, WindX['Ctrl_Pressed'])
            if key.vk == 67:
                if WindX['Text'] and (not WindX['stop_cheking']):
                    WindX['Text'].delete('1.0',END)
                    WindX['Status'].configure(text='',bg='#EFEFEF')

                if WindX['Ctrl_Pressed']:
                    WindX['Ctrl_C'] = 1
    except:  
        print(traceback.format_exc())      
        

def on_release(key):
    try:
        if isinstance(key, keyboard.Key):
            #print('Key:', key.name, ",", key.value.vk)
            if key.name == 'ctrl_l':
                WindX['Ctrl_Pressed'] = 0
                #print('special key {0} pressed'.format(key))

        if WindX['Ctrl_C']:
            CheckClipboard()
            WindX['Ctrl_C'] = 0

    except:  
        print(traceback.format_exc()) 

def keyboardListener():
    with keyboard.Listener(
        on_press=on_press,
        on_release=on_release) as listener:
        listener.join()

def MouseOnClick(x, y, button, pressed):
    #print(button,'{0} at {1}'.format('Pressed' if pressed else 'Released', (x, y)))
    #if re.match(r'.*left',str(button),re.I) and (not pressed):
    if WindX['TopLevel']:
        WindX['TopLevel'].destroy()
        WindX['TopLevel'] = None

def MouseListener():
    with mouseListener(on_click=MouseOnClick) as listener: #(on_move=on_move, on_click=on_click, on_scroll=on_scroll) 
        listener.join()

if __name__ == "__main__":     
    t1 = threading.Timer(1,keyboardListener)
    t2 = threading.Timer(1,main)
    t3 = threading.Timer(1,MouseListener)
    t1.start()  
    t2.start()  
    t3.start()  
    


'''
CAR Category .[CAR Category Workflow]
CAR Group .[CAR Group Workflow]
CAR Type .[CAR Type Workflow]
Business Sector .[Business Sector Workflow]
Audit Finding Type .[Audit Finding Type Workflow]
Audit Type .[Audit Type Workflow]
Segment .[Segment Workflow]
Severity .[Severity Workflow]
'''