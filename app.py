import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
import re

def main():
    st.set_page_config(
        page_title="残業時間集計アプリ",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 残業時間集計アプリ")
    st.markdown("---")
    
    # ファイルアップロード
    uploaded_file = st.file_uploader(
        "エクセルファイルをアップロードしてください",
        type=['xlsx', 'xls'],
        help="複数のシートを持つエクセルファイルをアップロードしてください"
    )
    
    if uploaded_file is not None:
        try:
            # エクセルファイルを読み込み
            workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
            sheet_names = workbook.sheetnames
            
            st.success(f"ファイルが正常に読み込まれました。シート数: {len(sheet_names)}")
            
            # 固定シートの確認
            fixed_sheets = ["まとめ", "記入例", "報告書format"]
            member_sheets = [sheet for sheet in sheet_names if sheet not in fixed_sheets]
            
            st.info(f"固定シート: {fixed_sheets}")
            st.info(f"メンバーシート: {member_sheets}")
            
            if member_sheets:
                # 残業時間の集計
                overtime_data = extract_overtime_data(workbook, member_sheets)
                
                if overtime_data:
                    display_results(overtime_data)
                else:
                    st.warning("残業時間のデータが見つかりませんでした。")
            else:
                st.warning("メンバーのシートが見つかりませんでした。")
                
        except Exception as e:
            st.error(f"ファイルの読み込み中にエラーが発生しました: {str(e)}")

def extract_overtime_data(workbook, member_sheets):
    """残業時間データを抽出する"""
    overtime_data = {}
    
    # 時間帯の定義
    time_slots = {
        'K39': '休日時間帯の応動（09:00-18:00）',
        'O39': '平日・休日時間外の応動（18:00-22:00）',
        'S39': '平日・休日深夜の応動（22:00-05:00）',
        'W39': '平日・休日時間外の応動（05:00-09:00）'
    }
    
    # デバッグ情報を表示
    st.markdown("### 🔍 デバッグ情報")
    
    for sheet_name in member_sheets:
        try:
            worksheet = workbook[sheet_name]
            member_data = {}
            
            # シートの基本情報を表示
            st.write(f"**シート: {sheet_name}**")
            
            for cell_ref, time_slot in time_slots.items():
                # セルK39, O39, S39, W39の値を取得
                cell_value = worksheet[cell_ref].value
                
                # 結合セルの場合、下のセル（K40, O40, S40, W40）も確認
                if cell_value is None:
                    # 結合セルの下のセルを確認
                    next_cell_ref = cell_ref.replace('39', '40')
                    cell_value = worksheet[next_cell_ref].value
                    st.write(f"- 結合セル確認: {next_cell_ref}: {cell_value}")
                
                # デバッグ情報を表示
                st.write(f"- {cell_ref}: {cell_value} (型: {type(cell_value)})")
                
                if cell_value is not None:
                    # 時間の形式をパース（例: "1:00" -> 1.0時間）
                    time_hours = parse_time_to_hours(str(cell_value))
                    st.write(f"  → パース結果: {time_hours}時間")
                    if time_hours > 0:
                        member_data[time_slot] = time_hours
                else:
                    member_data[time_slot] = 0
                    st.write(f"  → セルが空またはNone")
            
            # データがある場合のみ追加
            if any(value > 0 for value in member_data.values()):
                overtime_data[sheet_name] = member_data
                st.success(f"✅ {sheet_name}: データが見つかりました")
            else:
                st.warning(f"⚠️ {sheet_name}: データが見つかりませんでした")
                
        except Exception as e:
            st.error(f"❌ シート '{sheet_name}' の処理中にエラーが発生しました: {str(e)}")
            continue
    
    return overtime_data

def parse_time_to_hours(time_str):
    """時間文字列を時間数に変換する"""
    if not time_str or time_str.strip() == '':
        return 0
    
    # 文字列をクリーンアップ
    time_str = str(time_str).strip()
    
    # 時間:分の形式をパース（例: "1:30" -> 1.5時間）
    if ':' in time_str:
        try:
            parts = time_str.split(':')
            if len(parts) == 2:
                hours = int(parts[0])
                minutes = int(parts[1])
                return hours + minutes / 60
        except:
            pass
    
    # 数値のみの場合はそのまま返す
    try:
        return float(time_str)
    except:
        # 文字列から数値を抽出（例: "1時間30分" -> 1.5）
        import re
        numbers = re.findall(r'\d+\.?\d*', time_str)
        if numbers:
            return float(numbers[0])
        return 0

def display_results(overtime_data):
    """結果を表示する"""
    st.markdown("## 📈 残業時間集計結果")
    
    # データフレームを作成
    df_data = []
    for member, data in overtime_data.items():
        row = {'メンバー': member}
        row.update(data)
        df_data.append(row)
    
    if df_data:
        df = pd.DataFrame(df_data)
        
        # 列の順序を指定
        columns_order = ['メンバー'] + list(df.columns[1:])
        df = df[columns_order]
        
        # 表示
        st.dataframe(df, use_container_width=True)
        
        # ダウンロードボタン
        csv = df.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="📥 CSVファイルとしてダウンロード",
            data=csv,
            file_name=f"残業時間集計_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )
        
        # 統計情報
        st.markdown("### 📊 統計情報")
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("対象メンバー数", len(overtime_data))
        
        with col2:
            total_hours = sum(sum(data.values()) for data in overtime_data.values())
            st.metric("総残業時間", f"{total_hours:.1f}時間")
        
        with col3:
            avg_hours = total_hours / len(overtime_data) if overtime_data else 0
            st.metric("平均残業時間", f"{avg_hours:.1f}時間")
        
        with col4:
            max_hours = max(sum(data.values()) for data in overtime_data.values()) if overtime_data else 0
            st.metric("最大残業時間", f"{max_hours:.1f}時間")

if __name__ == "__main__":
    main()
