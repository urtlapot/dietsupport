import pandas as pd
import matplotlib.pyplot as plt
import tkinter
from tkinter import messagebox
import os



def create_user():
    global input_new_user_name
    global create_new_user

    create_new_user = tkinter.Toplevel()
    #print(type(create_new_user))
    create_new_user.title("new_user")
    create_new_user.geometry("240x60")
    #ウィンドウボックスをタイトルを設定
    #"x"は小文字のエックス

    new_user_label = tkinter.Label(create_new_user, text='登録者名')
    new_user_label.grid(row=1, column=1, padx=10,)
    #名前表記内容+表示位置設定

    input_new_user_name = tkinter.Entry(create_new_user, width = 20)
    input_new_user_name.insert(tkinter.END,"名前を入力してください。")
    input_new_user_name.grid(row=1, column=2)
    #入力ボックスを幅決めて作成+表示位置設定

    new_user_ok = tkinter.Button(create_new_user, text="決定",command=new_ok_click)
    new_user_ok.place(x=10,y=25)

    create_new_user.protocol("WM_DELETE_WINDOW", new_on_closing)
    create_new_user.mainloop()

def new_ok_click():
#決定実行時の処理
    global new_user_name
    #変数を関数外から呼び出せるように（非推奨）
    new_user_name = input_new_user_name.get()
    excel_url = str(os.path.dirname(__file__)) + str('/list/') + new_user_name + str('.xlsx')
    
    if os.path.exists(excel_url) == True:
        new_user_name = ''
        new_pop = tkinter.Tk()
        new_pop.withdraw()
        #空白のwindowが表示されるので、windowを作成→非表示に設定
        messagebox.showinfo('Check' ,'すでに登録されているユーザー名です。')
        #tkwindow表示されるの消す
        new_pop.destroy()          

    elif new_user_name == '':
        new_pop = tkinter.Tk()
        new_pop.withdraw()
        #空白のwindowが表示されるので、windowを作成→非表示に設定
        messagebox.showinfo('Check' ,'名前を入力してください。')
        #tkwindow表示されるの消す
        new_pop.destroy()     

    elif os.path.exists(excel_url) == False:
        new_df = pd.read_excel(str(os.path.dirname(__file__)) + str('/list/template.xlsx'))
        #print(new_df)
        new_df.to_excel(str(os.path.dirname(__file__)) + str('/list/') + new_user_name + str('.xlsx'), index=False)

        new_pop = tkinter.Tk()
        new_pop.withdraw()
        #空白のwindowが表示されるので、windowを作成→非表示に設定
        messagebox.showinfo('Check' ,'登録できました。')
        #tkwindow表示されるの消す
        new_pop.destroy() 
        create_new_user.destroy()        

    #リスト以外なら作成実行

def new_on_closing():
#ウィンドウ閉じたときの処理
    if messagebox.askokcancel("Quit", "新規登録を取りやめますか？"):
       create_new_user.destroy()   