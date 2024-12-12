import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
from datetime import datetime
import os

class DMSenderApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("DM送信数入力")
        self.csv_path = None
        self.setup_ui()
        
    def setup_ui(self):
        # CSVファイル選択ボタン
        tk.Button(self.root, text="CSVファイルを選択", command=self.select_excel).pack(pady=10)
        
        # 選択されたファイルパスを表示するラベル
        self.file_label = tk.Label(self.root, text="ファイルが選択されていません", wraplength=300)
        self.file_label.pack(pady=5)
        
        # 入力フィールドの作成
        tk.Label(self.root, text="送信するDM数を入力してください:").pack(pady=10)
        self.count_entry = tk.Entry(self.root)
        self.count_entry.pack(pady=5)
        
        # 実行ボタンの作成
        self.execute_button = tk.Button(self.root, text="実行", command=self.process_dm, state='disabled')
        self.execute_button.pack(pady=10)
        
    def select_excel(self):
        file_path = filedialog.askopenfilename(
            title="CSVファイルを選択",
            filetypes=[("CSVファイル", "*.csv")]
        )
        if file_path:
            self.csv_path = file_path
            self.file_label.config(text=f"選択されたファイル:\n{os.path.basename(file_path)}")
            self.execute_button.config(state='normal')  # ファイル選択後に実行ボタンを有効化
        
    def process_dm(self):
        if not self.csv_path:
            messagebox.showerror("エラー", "CSVファイルを選択してください")
            return
            
        try:
            print("1. 入力値の確認")
            count = int(self.count_entry.get())
            if count <= 0:
                messagebox.showerror("エラー", "正の整数を入力してください")
                return
                
            print("2. CSVファイル読み込み開始")
            # 複数のエンコーディングを試す
            encodings = ['cp932', 'shift_jisx0213', 'shift-jis', 'utf-8', 'utf-8-sig']
            for encoding in encodings:
                try:
                    print(f"エンコーディング {encoding} で試行中...")
                    df = pd.read_csv(self.csv_path, encoding=encoding)
                    print(f"成功: {encoding} で読み込みました")
                    break
                except Exception as e:
                    print(f"失敗: {encoding} - {str(e)}")
                    if encoding == encodings[-1]:
                        raise Exception("すべてのエンコーディングで読み込みに失敗しました")
                    continue
                
            print("3. CSVファイル読み込み完了")
            
            print("4. DM停止フラグ確認")
            df = df[df['ＤＭ停止'].fillna('').str.strip() != '停止']
            
            print("5. 発送回数のソート開始")
            latest_send_date = self._get_latest_send_date_column(df)
            df['送信回数'] = df[[col for col in df.columns if '回発送日' in col]].notna().sum(axis=1)
            df = df.sort_values('送信回数')
            print("6. ソート完了")
            
            print("7. 送信リスト作成開始")
            send_list = df.head(count)[['郵便番号', '宛先住所１', '送付先名', '管理者氏名']]
            send_list['敬称'] = '様'
            
            print("8. ファイル出力準備")
            excel_dir = os.path.dirname(self.csv_path)
            csv_filename = os.path.join(excel_dir, f'送信リ��ト_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv')
            
            print("9. CSVファイル出力")
            send_list.to_csv(csv_filename, index=False, encoding='cp932')
            
            print("10. 発送日更新処理開始")
            self._update_send_dates(df, send_list.index, latest_send_date)
            
            print("11. 処理完了")
            messagebox.showinfo("完了", f"処理が完了しました\n作成ファイル: {os.path.basename(csv_filename)}")
            self.root.destroy()
            
        except ValueError as ve:
            print(f"ValueError発生: {str(ve)}")
            messagebox.showerror("エラー", "有効な数値を入力してください")
        except Exception as e:
            print(f"エラー発生箇所の詳細: {str(e)}")
            print(f"エラーの種類: {type(e)}")
            messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{str(e)}")
            
    def _get_latest_send_date_column(self, df):
        send_date_cols = [col for col in df.columns if '回発送日' in col]
        max_filled_col = None
        max_filled_num = 8  # 第9回から始めるために8を初期値とする

        # 各行で最後に日付が入っている列を見つける
        for col in send_date_cols:
            try:
                num = int(col.replace('第', '').replace('回発送日', ''))
                if num > max_filled_num and not df[col].isna().all():
                    max_filled_num = num
                    max_filled_col = col
            except ValueError:
                continue

        # 次の列名を生成
        next_num = max_filled_num + 1 if max_filled_num >= 8 else 9
        new_col = f'第{next_num}回発送日'
        
        # 新しい列を適切な位置に挿入
        if new_col not in df.columns:
            # 最後の発送日列のインデックスを見つける
            last_date_col_idx = -1
            for i, col in enumerate(df.columns):
                if '回発送日' in col:
                    last_date_col_idx = i
            
            # 既存の列リストを取得
            cols = df.columns.tolist()
            
            # 新しい列を最後の発送日列の直後に挿入
            if last_date_col_idx != -1:
                cols.insert(last_date_col_idx + 1, new_col)
                df = df.reindex(columns=cols)
                df[new_col] = None
        
        return new_col
        
    def _update_send_dates(self, df, target_indices, date_column):
        try:
            # 元のCSVファイルを読み直す
            original_df = pd.read_csv(self.csv_path, encoding='cp932')
            
            # 発送日の列のみを更新
            today = datetime.now().strftime('%Y年%m月%d日')
            original_df.loc[target_indices, date_column] = today
            
            # 新しいファイル名を生成
            file_name = os.path.splitext(self.csv_path)[0]
            new_excel_path = f"{file_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            
            # 更新したデータを新しいExcelファイルとして保存
            original_df.to_excel(new_excel_path, index=False, engine='openpyxl')
            
        except Exception as e:
            print(f"更新処理でエラー発生: {str(e)}")
            raise
        
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = DMSenderApp()
    app.run() 