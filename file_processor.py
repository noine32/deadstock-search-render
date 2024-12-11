import pandas as pd
import chardet
import io
from datetime import datetime

class FileProcessor:
    @staticmethod
    def detect_encoding(file_bytes):
        result = chardet.detect(file_bytes)
        return result['encoding']

    @staticmethod
    def read_excel(file):
        try:
            df = pd.read_excel(file, engine='openpyxl')
            return df
        except Exception as e:
            raise Exception(f"Excelファイルの読み込みエラー: {str(e)}")

    @staticmethod
    def read_csv(file, file_type='default'):
        try:
            file_bytes = file.getvalue()
            encoding = FileProcessor.detect_encoding(file_bytes)
            
            if file_type == 'inventory':
                # 不良在庫データの場合、最初の7行をスキップ
                df = pd.read_csv(io.BytesIO(file_bytes), encoding=encoding, skiprows=7)
            else:
                df = pd.read_csv(io.BytesIO(file_bytes), encoding=encoding)
            
            return df
        except Exception as e:
            raise Exception(f"CSVファイルの読み込みエラー: {str(e)}")

    @staticmethod
    def process_data(purchase_df, inventory_df, yj_code_df):
        print("処理開始: データフレームの初期状態")
        print(f"Inventory shape: {inventory_df.shape}")
        print(f"Purchase history shape: {purchase_df.shape}")
        print(f"YJ code shape: {yj_code_df.shape}")
        
        # 空の薬品名を持つ行を削除
        inventory_df = inventory_df[inventory_df['薬品名'].notna() & (inventory_df['薬品名'] != '')].copy()
        print("空の薬品名を削除後の inventory shape:", inventory_df.shape)
        
        # データフレームの型チェックとNaN値の処理
        for df_name, df in {"inventory": inventory_df, "purchase_history": purchase_df, "yj_code": yj_code_df}.items():
            if not isinstance(df, pd.DataFrame):
                print(f"Warning: {df_name} is not a DataFrame")
                continue
            for col in df.columns:
                df[col] = df[col].fillna('').astype(str)
        
        # 在庫金額CSVから薬品名とＹＪコードのマッピングを作成
        if not yj_code_df.empty:
            yj_mapping = dict(zip(yj_code_df['薬品名'], zip(yj_code_df['ＹＪコード'], yj_code_df['単位'])))
        else:
            print("Warning: YJ code DataFrame is empty")
            yj_mapping = {}
        
        # 不良在庫データに対してＹＪコードと単位を設定
        inventory_df['ＹＪコード'] = inventory_df['薬品名'].map(lambda x: yj_mapping.get(x, (None, None))[0])
        inventory_df['単位'] = inventory_df['薬品名'].map(lambda x: yj_mapping.get(x, (None, None))[1])
        
        # マージ前の状態を確認
        print("\nマージ前のデータ確認:")
        print("Inventory columns:", inventory_df.columns.tolist())
        print("Purchase history columns:", purchase_df.columns.tolist())
        
        # 必要なカラムが存在することを確認
        required_columns = ['厚労省CD', '法人名', '院所名', '品名・規格', '新薬品ｺｰﾄﾞ']
        missing_columns = [col for col in required_columns if col not in purchase_df.columns]
        
        if missing_columns:
            print(f"Warning: Missing columns in purchase_history_df: {missing_columns}")
            # 不足しているカラムを空の文字列で追加
            for col in missing_columns:
                purchase_df[col] = ''
        
        # ＹＪコードと厚労省CDで紐付け
        try:
            merged_df = pd.merge(
                inventory_df,
                purchase_df[required_columns],
                left_on='ＹＪコード',
                right_on='厚労省CD',
                how='left'
            )
            print("マージ後のデータ形状:", merged_df.shape)
        except Exception as e:
            print(f"マージ中にエラーが発生: {str(e)}")
            return pd.DataFrame()
        
        # 院所名別にデータを整理
        result_df = merged_df[[
            '品名・規格',
            '在庫量',
            '単位',
            '新薬品ｺｰﾄﾞ',
            '使用期限',
            'ロット番号',
            '法人名',
            '院所名'
        ]].copy()
        
        # 空の値を空文字列に変換
        result_df = result_df.fillna('')
        
        # 空の品名・規格を持つ行を削除
        result_df = result_df[result_df['品名・規格'].notna() & (result_df['品名・規格'] != '')].copy()
        
        # 院所名でソート
        result_df = result_df.sort_values(['法人名', '院所名'])
        
        return result_df

    @staticmethod
    def generate_excel(df):
        excel_buffer = io.BytesIO()

        try:
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                # 院所名ごとにシートを作成（空の値を除外）
                for name in df['院所名'].unique():
                    if not pd.isna(name) and str(name).strip():
                        sheet_name = str(name)[:31].strip()  # シート名を31文字以内に制限
                        sheet_df = df[df['院所名'] == name].copy()
                        
                        if not sheet_df.empty:
                            try:
                                # 法人名と院所名を取得
                                houjin_name = str(sheet_df['法人名'].iloc[0] or '')
                                insho_name = str(sheet_df['院所名'].iloc[0] or '')
                                
                                # 表示用のカラムから法人名と院所名を除外
                                display_df = sheet_df.drop(['法人名', '院所名'], axis=1)
                                
                                # 「引取り可能数」列を追加
                                display_df.insert(display_df.columns.get_loc('ロット番号') + 1, '引取り可能数', '')
                                
                                # ヘッダーデータの作成
                                header_data = pd.DataFrame([
                                    ['不良在庫引き取り依頼'],
                                    [''],
                                    [f"{houjin_name.strip()} {insho_name.strip()}", '', '御中'],
                                    [''],
                                    ['下記の不良在庫につきまして、引き取りのご検討を賜れますと幸いです。どうぞよろしくお願いいたします。'],
                                    ['']
                                ])
                                
                                # ヘッダーとデータを書き込み
                                header_data.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                                display_df.to_excel(writer, sheet_name=sheet_name, startrow=6, index=False)
                                
                                # シートを取得してフォーマットを設定
                                worksheet = writer.sheets[sheet_name]
                                
                                # 列幅の設定
                                worksheet.column_dimensions['A'].width = 35  # 品名・規格
                                worksheet.column_dimensions['B'].width = 15  # 在庫量
                                worksheet.column_dimensions['C'].width = 10  # 単位
                                worksheet.column_dimensions['D'].width = 15  # 新薬品コード
                                worksheet.column_dimensions['E'].width = 15  # 使用期限
                                worksheet.column_dimensions['F'].width = 15  # ロット番号
                                worksheet.column_dimensions['G'].width = 20  # 引取り可能数
                                
                                # 行の高さを設定（30ピクセル）
                                for row in range(1, worksheet.max_row + 1):
                                    worksheet.row_dimensions[row].height = 30
                                
                                # デフォルトのフォントサイズを14に設定
                                for row in worksheet.rows:
                                    for cell in row:
                                        if cell.font is None:
                                            cell.font = cell.font.copy(size=14)
                                        else:
                                            cell.font = cell.font.copy(size=14)
                                
                                # タイトルのフォント設定（サイズ16）
                                cell_a1 = worksheet['A1']
                                cell_a1.font = cell_a1.font.copy(size=16)
                                
                                # 法人名・院所名と御中のフォント設定（サイズ14、太字）
                                cell_a3 = worksheet['A3']  # 法人名・院所名
                                cell_c3 = worksheet['C3']  # 御中
                                font_style = cell_a3.font.copy(size=14, bold=True)
                                cell_a3.font = font_style
                                cell_c3.font = font_style
                                
                            except Exception as e:
                                print(f"シート '{sheet_name}' の処理中にエラーが発生: {str(e)}")
                                continue
        
        except Exception as e:
            print(f"Excel生成中にエラーが発生: {str(e)}")
            return None

        excel_buffer.seek(0)
        return excel_buffer
