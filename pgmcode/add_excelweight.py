import openpyxl
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import tkinter
from tkinter import messagebox
import datetime
import os


class AddData:
    def __init__(self, user_name, excel):
        self.user_name = user_name
        self.excel_url = excel

    def input_data(self):
        global input_add_date
        global input_add_weight
        global input_add_etc
        global add_data_form

        day = datetime.date.today()
        today = day.strftime("%Y/%m/%d")
        #PGM起動時の日付を取得→入力用に形式を変換してtodayに格納

        add_data_form = tkinter.Toplevel()
        #toplevelで出すことでメニュー画面を残しておく+何度でも展開可能（重複してしまう対策必要あり）

        add_data_form.title("add_data_form")
        add_data_form.geometry("220x100")
        
        #ウィンドウボックスをタイトルを設定
        #"x"は小文字のエックス

        add_date_label = tkinter.Label(add_data_form, text='測定日')
        add_date_label.grid(row=1, column=1, padx=10,)
        #tkinter設定時にadd_data_formを指定する必要あり
        #名前表記内容+表示位置設定
        input_add_date = tkinter.Entry(add_data_form, width = 20)
        input_add_date.insert(tkinter.END,today)
        input_add_date.grid(row=1, column=2)
        #入力ボックスを幅決めて作成+表示位置設定

        add_weight_label = tkinter.Label(add_data_form, text='体重')
        add_weight_label.grid(row=2, column=1, padx=10,)
        #tkinter設定時にadd_data_formを指定する必要あり
        #名前表記内容+表示位置設定
        input_add_weight = tkinter.Entry(add_data_form, width = 20)
        input_add_weight.insert(tkinter.END,"")
        input_add_weight.grid(row=2, column=2)
        #入力ボックスを幅決めて作成+表示位置設定

        add_etc_label = tkinter.Label(add_data_form, text='備考')
        add_etc_label.grid(row=3, column=1, padx=10,)
        #tkinter設定時にadd_data_formを指定する必要あり
        #名前表記内容+表示位置設定
        input_add_etc = tkinter.Entry(add_data_form, width = 20)
        input_add_etc.insert(tkinter.END,"")
        input_add_etc.grid(row=3, column=2)
        #入力ボックスを幅決めて作成+表示位置設定

 

        user_ok = tkinter.Button(add_data_form, text="決定",command=self.ok_click_add)
        user_ok.place(x=150,y=70)
        #tkinter設定時にadd_data_formを指定する必要あり

        add_data_form.protocol("WM_DELETE_WINDOW", self.on_closing_add)
        add_data_form.mainloop()


    def ok_click_add(self):
    #決定実行時の処理
        global add_date
        global add_weight
        global add_etc
        #変数を関数外から呼び出せるように（非推奨）

        add_date = input_add_date.get()
        add_weight = input_add_weight.get()
        add_etc = input_add_etc.get()

        if add_etc == '':
            messagebox.showinfo('Check' ,'測定日: ' + add_date + '\n' + '体重: ' + add_weight + '　が入力されました\n')
            #入力値をメッセージボックスに表示
            self.add_datas()
            add_data_form.destroy()
            #取得したデータをexcelに追加してからウィンドウを閉じる
        else:
            messagebox.showinfo('Check' ,'測定日: ' + add_date + '\n' + '体重: ' + add_weight + '\n' + '備考: ' + add_etc + '　が入力されました\n')
            #入力値をメッセージボックスに表示
            self.add_datas()
            add_data_form.destroy()
            #取得したデータをexcelに追加してからウィンドウを閉じる

    def on_closing_add(self):
    #ウィンドウ閉じたときの処理
        if messagebox.askokcancel("Quit", "入力を中止しますか?"):
            add_data_form.destroy()



    def add_datas(self):
            add_df = pd.read_excel(self.excel_url)
            add_df2 = pd.DataFrame([[add_date, add_weight, add_etc]],
                    index=[str(len(add_df))],columns=['測定日', '体重', '備考'])
            add_df3 = add_df.append(add_df2)  
            add_df3.to_excel(self.excel_url, index=False)
            #print(add_df3)
            


#####################################   修正用PGM　開始   ##############################################

    def delete_data(self):
        global input_del_date
        global delete_form
        day = datetime.date.today()
        today = day.strftime("%Y/%m/%d")

        delete_form = tkinter.Toplevel()
        delete_form.title("delete_form")
        delete_form.geometry("240x60")
        #toplevelで出すことでメニュー画面を残しておく+何度でも展開可能（重複してしまう対策必要あり）
        #ウィンドウボックスをタイトルを設定
        #"x"は小文字のエックス

        del_date_label = tkinter.Label(delete_form, text='削除する日付')
        del_date_label.grid(row=1, column=1, padx=10,)
        #tkinter設定時にdelete_formを指定する必要あり
        #名前表記内容+表示位置設定
        input_del_date = tkinter.Entry(delete_form, width = 20)
        input_del_date.insert(tkinter.END,today)
        input_del_date.grid(row=1, column=2)
        #入力ボックスを幅決めて作成+表示位置設定

        del_ok = tkinter.Button(delete_form, text="確認",command=self.ok_click_del)
        del_ok.place(x=180,y=20)
        #tkinter設定時にdelete_formを指定する必要あり

        delete_form.protocol("WM_DELETE_WINDOW", self.on_closing_del)
        delete_form.mainloop()      
        


    def del_datas(self):

        delete_dates = []
        delete_weight = []
        hit_counts = 0
        weight = ''

        del_df = pd.read_excel(self.excel_url)
        check_date = del_df['測定日'] == del_date
        #エクセルの呼び出し→測定日の中で入力値と一致するものをbool値で格納

        for ch in check_date:
            if ch == True:
                hit_counts += 1
        #入力値と一致するものの個数をカウント

        if hit_counts > 0:
        #一つ以上一致する時の処理
            hit_list = del_df.index[del_df['測定日'] == del_date]
            for i in hit_list:
                if i != 'None':
                    delete_weight.append(del_df.loc[i,'体重'])
            #入力した測定日の体重をリスト形式で取得

            for m in delete_weight:
                weight += str(m) + 'kg' + '\n'
            #取得したリストからMsgBox用に文字列を変換

            res = messagebox.askyesno('Coution' ,str(del_date) + ' の測定結果'\
                                            + '\n' + weight\
                                            + '以上、' + str(hit_counts) +'個のデータを削除しますか？')

            if res == True:
                del_df.drop(del_df.index[del_df['測定日'] == del_date],inplace = True)
                del_df.to_excel(self.excel_url, index=False)
                messagebox.showinfo('Check' ,'削除しました。')
                delete_form.destroy()
                #入力日のデータを削除してエクセルに上書き

            elif res == False:
                messagebox.showinfo('Check' ,'キャンセルしました。')
                #delete_form.destroy()

        elif hit_counts == 0:
        #一つも該当しない時の処理    
            messagebox.showinfo('Coution' ,'入力日の測定データはありません。\n'\
                                            + 'もう一度入力してください。')
            #delete_form.destroy()


    def ok_click_del(self):
    #決定実行時の処理
        global del_date
        global del_df
        del_date = input_del_date.get()

        if del_date != '':
            self.del_datas()

        elif del_date == '':
            messagebox.showinfo('Check' ,'削除する日付を入力してください。')
            #入力値をメッセージボックスに表示



    def on_closing_del(self):
    #ウィンドウ閉じたときの処理
        if messagebox.askokcancel("Quit", "入力を中止しますか?"):
            delete_form.destroy()

#####################################   修正用PGM　終了   ##############################################



if __name__ == '__main__':
    user_name = 'for_test'
    url = str(os.path.dirname(__file__)) + str('/list/') + user_name + str('.xlsx')
    dele = AddData(user_name, url)
    dele.delete_data()
