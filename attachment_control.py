#Credit - https://stackoverflow.com/questions/40388219/typeerror-str-object-is-not-callable-with-win32com-interfacing-with-attachmat

import win32com.client
import time

system = win32com.client.Dispatch("EXTRA.System")
sess0 = system.ActiveSession
screen = sess0.Screen

commit_id = []
loan_number = []
string_to_look_up = list(zip(commit_id, loan_number))

def write(screen,row,col,text):
    screen.row = row
    screen.col = col
    screen.SendKeys(text)
    while screen.OIA.XStatus != 0: pass

def read(screen,row,col,length,page=None):
    if page is None:
        return screen.Area(row, col, row, col+length).value
    else:
        return screen.Area(row, col, row, col+length, page).value

def navigate_pages(order_number):
    while screen.OIA.XStatus != 0: pass
    write(screen, 22, 8, order_number)
    screen.Sendkeys("<Enter>")
    while screen.OIA.XStatus != 0: pass

def commit_navigate_to_maint_menu():
    commitment = ['EMNA', 'ECMM', 'ECMN', 'ECSM', 'ECSS', 'ECSL'] #ECSM is the main menu for commitment
    short_cut_to = [0, 1, 2, 3, 4, 4] #how far (number of pages) from EMNA
    dic_commitment = dict(zip(commitment, short_cut_to))
    current_instance = str(read(screen, 1,1,6).strip())
    where_to = 'ECMN'
    if current_instance in commitment:
        get_number = dic_commitment[current_instance] - commitment.index(where_to.strip())
        if get_number >=0 :
            for i in range(0, get_number):
                while screen.OIA.XStatus != 0: pass
                screen.Sendkeys("<pF12>")
                while screen.OIA.XStatus != 0: pass
        elif get_number < 0:
            get_number =  get_number * -1
            # print(get_number)
            for i in range(0, get_number):
                navigate_pages(1)
                while screen.OIA.XStatus != 0: pass
    elif current_instance not in commitment:
        screen.Sendkeys("<pF3>")
        while screen.OIA.XStatus != 0: pass
        for i in range(0, 2):
            navigate_pages(1)
            while screen.OIA.XStatus != 0: pass
    else:
        print('error - I suggest that you go to the very main page, and re-run the file')
        raise SystemExit

def main():
    commit_navigate_to_maint_menu()
    while screen.OIA.XStatus != 0: pass
    for i,x in string_to_look_up:
        write(screen, 12, 46, i)
        write(screen, 15, 46, x)
        navigate_pages('s')
        navigate_pages(1)
        if read(screen,6, 20, 50).strip() == 'APPLICATION APPROVED':
            print('approved')
            screen.Sendkeys("<pF3>")
            while screen.OIA.XStatus != 0: pass
            commit_navigate_to_maint_menu()
            # navigate_pages('5')
        else:
            print('not approved')
            screen.Sendkeys("<pF3>")
            while screen.OIA.XStatus != 0: pass
            commit_navigate_to_maint_menu()

main()


