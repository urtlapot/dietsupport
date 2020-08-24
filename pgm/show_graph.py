import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import tkinter
from tkinter import messagebox
import datetime
import os


class MakeGraph:
    def __init__(self,user_name, excel):
        self.user_name = user_name
        self.excel_url = excel
        
    def draw_graph(self):
        df = pd.read_excel(self.excel_url)
        df = df.sort_values(['測定日','体重'])
        df = df.reset_index()
        df = df.drop([0])
        # ラベル毎の値を取り出し
        #sheetname指定したらkeyエラーが発生

        date = df['測定日']
        #df.loc['1':,['測定日']]
        weight = df['体重']
        #df.loc['1':,['体重']]

        # グラフ化
        plt.figure(figsize=(9,6))
        plt.title(self.user_name + '\'s weight')
        plt.xlabel('measuring date')
        plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%m/%d'))
        plt.ylabel('weight')
        plt.plot(date, weight)
        plt.grid()
        plt.show()
        # ワークブックを取得する
        ## 要変更！！excel指定or自動選択に！ ##

    def comparison_graph(self):
        all_files = []
        all_user = []
        all_user_df = pd.read_excel(str(os.path.dirname(__file__)) + str('/list/template.xlsx'))
        #ひな形としてtemplateファイル読み込み

        users_url = str(os.path.dirname(__file__)) + str('/list/')
        all_files = os.listdir(str(os.path.dirname(__file__)) + str('/list'))


        for n in all_files:
            if n[-5:] == '.xlsx':
                all_user.append(n)
        all_user.remove('template.xlsx')
        #全xlsxファイル保存先からexcelファイルのみを取り出し
        #templateをリストから削除



        plt.figure(figsize=(9,6))
        plt.title('all user\'s weight')
        plt.xlabel('measuring date')
        plt.ylabel('weight')
        plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%m/%d'))
        #グラフの外枠の大きさを設定
        #X/Yのラベル名を決定
        #X軸の日付をmm/dd表記に変更

        # グラフ化
        for u in all_user:
            df2 = pd.read_excel(users_url + u)
            df2 = df2.sort_values(['測定日','体重'])
            df2 = df2.reset_index()
            df2 = df2.drop([0])
            df2 = df2.rename(columns={'体重': u[:-5]})
            #df2へ一時的に各ユーザーのdfを追加　→'体重'を'ユーザー名'に変更
            #print(df2)
            all_user_df = pd.concat([all_user_df,df2])
            #all_user_dfにデータ追加
            plt.plot('測定日', u[:-5], data=all_user_df)
            #折れ線グラフとして読み込み

        plt.legend(bbox_to_anchor=(0,1), loc='upper left', borderaxespad=0, fontsize=18)
        #凡例を表示　グラフに対する位置座標,座標を合わせる位置,外枠と合わせ角の距離,文字サイズ
        plt.grid()
        plt.show()
        



