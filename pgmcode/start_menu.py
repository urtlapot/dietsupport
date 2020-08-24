import pandas as pd
import matplotlib.pyplot as plt
import os
import os.path
import tkinter
from tkinter import messagebox
import sys

import show_graph
import add_excelweight
import create_new


#############################        user select start        #############################
def select_user_window():
    global input_user_name

    select_user.title("select_user")
    select_user.geometry("260x80")
    #ウィンドウボックスをタイトルを設定
    #"x"は小文字のエックス

    input_user_name_label = tkinter.Label(text='使用者名')
    input_user_name_label.grid(row=1, column=1, padx=10,)
    #名前表記内容+表示位置設定

    input_user_name = tkinter.Entry(width = 20)
    input_user_name.insert(tkinter.END,"名前を入力してください。")
    input_user_name.grid(row=1, column=2)
    #入力ボックスを幅決めて作成+表示位置設定

    user_ok = tkinter.Button(text="決定",command=ok_click)
    user_ok.place(x=10,y=25)

    new_user = tkinter.Button(text="新規ユーザー",command=create)
    new_user.place(x=50,y=25)

    select_user.protocol("WM_DELETE_WINDOW", on_closing)


def create():
    create_new.create_user()


def ok_click():
#決定実行時の処理
    global user_name
    #変数を関数外から呼び出せるように（非推奨）
    user_name = input_user_name.get()


    if user_name == '' or user_name == '名前を入力してください。':
        messagebox.showinfo('Caution','ユーザー名を入力してください')
        
    elif user_name == 'template':
        user_name = ''
        messagebox.showinfo('Check' ,'システム用のファイル名の為無効です。')

    else:
        messagebox.showinfo('Check' ,'ユーザー名: ' + user_name + '　が選択されました')
        #入力値をメッセージボックスに表示
        select_user.destroy()
        #ウィンドウを閉じる

def on_closing():
#ウィンドウ閉じたときの処理
    if messagebox.askokcancel("Quit", "アプリを終了しますか?"):
       select_user.destroy()   
#############################        user select end        #############################





#############################        menu select start        #############################
def select_menu_window():

    select_menu.title("select_menu")
    select_menu.geometry("300x240")
    #ウィンドウボックスをタイトルを設定
    #"x"は小文字のエックス

    show_user_name_label = tkinter.Label(text='ようこそ  ' + user_name + '  さん')
    show_user_name_label.grid(row=0, column=1, padx=10)
    #名前表記内容+表示位置設定



    input_weight = tkinter.Button(text="体重入力",command=add_data, height=2,width=15)
    input_weight.place(x=20,y=50)
    #input_weight.grid(rowspan=5, row=2, columnspan=3, column=2, padx=10)

    fix_weight = tkinter.Button(text="グラフ描画",command=drowing_graph, height=2,width=15)
    fix_weight.place(x=160,y=50)
    #fix_weight.grid(row=2, column=5, padx=10)
    


    drow_graph = tkinter.Button(text="入力値削除",command=erase_data, height=2,width=15)
    drow_graph.place(x=20,y=100)
    #drow_graph.grid(row=7, column=2, padx=10)

    comparision_user = tkinter.Button(text="ユーザー推移比較",command=com_graph, height=2,width=15)
    comparision_user.place(x=160,y=100)
    #comparision_user.grid(row=7, column=5, padx=10)


    close_menu = tkinter.Button(text="終了",command=on_closing_menu, height=2,width=15)
    close_menu.place(x=160,y=150)
    #close_menu.grid(row=12, column=5, padx=10)

    select_menu.protocol("WM_DELETE_WINDOW", on_closing_menu)


def on_closing_menu():
#ウィンドウ閉じたときの処理
    if messagebox.askokcancel("Quit", "アプリを終了しますか?"):
        select_menu.destroy() 


def test_click():
#決定実行時の処理
    messagebox.showinfo('test','未実装の機能です')
#############################        menu select end        #############################



#############################        menu action start        #############################
def drowing_graph():
    graph = show_graph.MakeGraph(user_name, excel_url)
    #入力したユーザー名をグラフ関数に渡す
    graph.draw_graph()

def com_graph():
    graph = show_graph.MakeGraph(user_name, excel_url)
    graph.comparison_graph()

def add_data():

    data = add_excelweight.AddData(user_name, excel_url)
    data.input_data()

def erase_data():

    era_data = add_excelweight.AddData(user_name, excel_url)
    era_data.delete_data()   

    
#############################        menu action end        #############################


## setting value ##
user_name = ''
#エラー対策の為だけに残している
#ユーザー入力画面で選択できるようする機能追加するかも


######################## start pgm ########################
if __name__ == '__main__':
    select_user = tkinter.Tk()
    select_user_window()
    select_user.mainloop()
    #入力画面設定用
    #user_nameの取得 or 新規userの登録



    excel_url = str(os.path.dirname(__file__)) + str('/list/') + user_name + str('.xlsx')
    #pyファイル内に'list'フォルダ→excelファイルで入れていると自動で開く
    #user_name取得後にexcelのurlを取得->メニューから動作選択した時にurlを渡す


    if user_name != '' and os.path.exists(excel_url) == True:
        #フォルダからuser_nameと一致するexcelを取得して不一致ならはじく
        select_menu = tkinter.Tk()
        select_menu_window()
        select_menu.mainloop()
        #メニュー画面展開用

    elif user_name != '':
        pop = tkinter.Tk()
        pop.withdraw()
        #空白のwindowが表示されるので、windowを作成→非表示に設定
        messagebox.showinfo('Check' ,'登録されていないユーザー名です。\n' + 'プログラムを終了します。')
        #tkwindow表示されるの消す
        pop.destroy()

