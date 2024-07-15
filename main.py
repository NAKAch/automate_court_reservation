from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
from selenium.common.exceptions import NoSuchElementException
import streamlit as st
import pandas as pd
from st_aggrid import AgGrid as ag

class CourtReserveBot:
    def __init__(self, url):
        self.url = url

    def landing_page(self):
        st.markdown("""# コート予約自動化ツール

## 説明
このツールは京都市のコート抽選予約を自動化するためのツールです。
                    
## 使い方
### 1. メニューバーを開き、「情報入力」から予約情報を入力
- 予約1～予約10までのタブを使って、予約情報を入力します。(全てのタブを使う必要はありません。)
- コートカードのID、パスワード、予約するコート、予約する日付、予約する時間帯を入力します。
- 「情報を保存する」をクリックして、入力情報を保存します。
                    
### 2. メニューバーを開き、「入力確認・予約実行」をクリック
- 予約情報の確認、削除、予約実行を行います。
- 予約情報では、保存している予約情報を確認できます。
- 予約情報の削除では、予約情報を削除できます。選択したのち、「削除する」をクリックして削除します。
- 「情報の更新」をクリックすると、「予約情報」が更新されます。
- 「予約を実行する」をクリックすると、保存している予約情報を元に、予約を実行します。
                    
### 3. その他
- 予約が成功or失敗した場合、「成功」or「失敗」と表示されます。
- メニューの「情報入力」から他のメニューに移動した場合、入力した情報はテキストボックスから消えます。\nですが、ちゃんと保存されているので安心してください。「入力確認・予約実行」から、保存情報を確認できます。
                    
## 注意事項
- このツールは京都市のコート抽選予約のみに対応しています。
- このツールを使用する際は、自己責任でお願いします。ツールの作成者は、一切の責任を取りません。
- 岡崎、宝の2つにのみ対応しています。
- 予約する時間帯は8～10、4～6、6～9の3つのみ対応しています。
""")

    def front(self,tab_name):
        st.markdown("## 情報の入力")
        st.markdown("### 使用するコートカードを入力して入力してください。")
        id = st.text_input(label="コートカードのIDを入力してください。", key=f"court_id_{tab_name}").replace(" ", "").replace("　", "")
        password = st.text_input(label="コートカードのパスワードを入力してください。", key=f"pass_{tab_name}").replace(" ", "").replace("　", "")
        st.markdown("### 予約するコートを選択してください。")
        park_options = ["岡崎", "宝"]
        park = st.selectbox(label="予約するコートを選択してください。",options=park_options,key=f"park_{tab_name}")
        st.markdown("### 予約する日付を選択してください。")
        date_options = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31]
        selected_date = st.multiselect(label="日付を選択してください。",
            options=date_options,
            max_selections=10,
            placeholder="日付を選択してください。",
            format_func=lambda x: f"{x}日",
            label_visibility='collapsed',
            key=f"date_{tab_name}"
        )
        selected_date = sorted(selected_date)
        if len(selected_date) < 10:
            st.write(f"あと{10-len(selected_date)}個の日付を選択できます。")
            st.write(f"選択された日付: {selected_date}")
        else:
            st.write(f"選択された日付: {selected_date}")
            st.write("これ以上日付を選択できません。")
        st.write("日付が確定しました。")
        st.markdown("### 予約する時間帯を選択してください。")
        reservation_lots = []
        time_options = ["未選択","8~10", "4~6", "6~9"]
        st.markdown("#### 一括入力")
        st.write("一括入力を選択すると、選択した全ての日付に対して、同じ時間帯を一括でセットします。")
        tmp_bulk_time = st.selectbox(label="時間帯を選択・修正してください。",options=time_options,index=0,key=f"bulk_time_{tab_name}")
        st.markdown("#### 個別入力・一括入力の個別修正")
        for i in selected_date:
            lot = []
            selected_time = st.selectbox(label=f"{i}日の時間帯を選択してください。",options=time_options,index=time_options.index(tmp_bulk_time),key=f"{i}_time_{tab_name}")
            lot.append(str(i))
            lot.append(selected_time)
            reservation_lots.append(lot)
        # 入力データを保存
        user_input = {
            "lot": reservation_lots,
            "park": park,
            "id": id,
            "password": password,
        }
        if len(user_input["lot"])>=1:
            for one_lot in user_input["lot"]:
                is_blank = any(i == "未選択" for i in one_lot) or user_input["park"] == "" or user_input["id"] == "" or user_input["password"] == ""
        else:
            is_blank = True
        save_button = st.button("情報を保存する",key=f"submit_{tab_name}")
        if is_blank:
            st.write("未入力の項目があります。")
            st.write("全ての項目を入力してください。")
        if save_button and is_blank:
            st.write("未入力の項目があります。")
            st.write("全ての項目を入力してください。")
        if save_button and not is_blank:
            st.write("情報を保存しました。")
            return user_input
    
    def edit_info(self,summery_list):
        filled_key_list = []
        arrays = []
        for key in summery_list.keys():
            if summery_list[key] != "":
                filled_key_list.append(key)
        for key in filled_key_list:
            tmp_array = []
            tmp_array.append(key)
            tmp_array.append(summery_list[key]["id"])
            tmp_array.append(summery_list[key]["password"])
            tmp_array.append(summery_list[key]["park"])
            for lot in summery_list[key]["lot"]:
                one_lot = f"{lot[0]}日 {lot[1]}"
                tmp_array.append(one_lot)
            for i in range(14-len(tmp_array)):
                tmp_array.append("")
            arrays.append(tmp_array)
        df = pd.DataFrame(arrays,columns=["Index","ID","パスワード","コート","予約日時1","予約日時2","予約日時3","予約日時4","予約日時5","予約日時6","予約日時7","予約日時8","予約日時9","予約日時10"])
        st.markdown("## 予約情報の確認")
        st.markdown("### 予約情報")
        st.write(df)
        st.markdown("### 予約情報の削除")
        selected_delete_index = st.multiselect(
            label="削除する予約を選択してください。",
            options=filled_key_list,
            max_selections=10,
            placeholder="削除する予約を選択してください。",
            format_func=lambda x: f"{x}",
            label_visibility='collapsed'
        )
        if st.button("削除する"):
            for index in selected_delete_index:
                summery_list[index] = ""
        st.button("情報の更新")
        pass

    def drive_website(self, lots, park, user_id, password):
        self.driver = webdriver.Chrome()
        self.wait = WebDriverWait(self.driver, 10)
        self.driver.get(self.url)
        time.sleep(3)
        response = True
        if response:
            confirm = True
            if confirm:
                try:
                    self.driver.switch_to.frame('MainFrame')
                    for lot in lots:
                        time.sleep(5)
                        self.driver.find_element(By.LINK_TEXT, 'テニス').click()
                        self.driver.find_element(By.LINK_TEXT, '京都市左京区').click()
                        self.driver.find_element(By.NAME, 'btn_next').click()
                        if park == '岡崎':
                            self.driver.find_element(By.XPATH, '/html/body/form/div[2]/center/table[4]/tbody/tr/td/table/tbody/tr[1]/td[3]/input').click()
                        elif park == '宝':
                            self.driver.find_element(By.XPATH, '/html/body/form/div[2]/center/table[4]/tbody/tr/td/table/tbody/tr[3]/td[7]/input').click()
                        self.driver.find_element(By.XPATH, '/html/body/form/div[2]/div[1]/center/table[4]/tbody/tr[1]/th[3]/a').click()
                        
                        self.driver.find_element(By.LINK_TEXT, str(lot[0])).click()
                        time_xpath = [
                            self.driver.find_element(By.XPATH, '/html/body/form/div[2]/div[2]/left/table[3]/tbody/tr/td/table/tbody/tr[5]/td[1]/a/img'),
                            self.driver.find_element(By.XPATH, '/html/body/form/div[2]/div[2]/left/table[3]/tbody/tr/td/table/tbody/tr[5]/td[5]/a/img'),
                            self.driver.find_element(By.XPATH, '/html/body/form/div[2]/div[2]/left/table[3]/tbody/tr/td/table/tbody/tr[5]/td[6]/a/img')
                        ]
                        if lot[1] == '8~10':
                            time_xpath[0].click()
                        elif lot[1] == '4~6':
                            time_xpath[1].click()
                        elif lot[1] == '6~9':
                            time_xpath[2].click()

                        original_window = self.driver.current_window_handle
                        new_window_handle = WebDriverWait(self.driver, 10).until(
                            EC.new_window_is_opened(self.driver.window_handles)
                        )
                        self.driver.switch_to.window(new_window_handle)
                        
                        select_element = Select(self.driver.find_element(By.TAG_NAME, 'select'))
                        select_element.select_by_value('1')
                        self.driver.find_element(By.XPATH, '/html/body/form/p/input[1]').click()
                        self.driver.switch_to.window(original_window)
                        self.driver.switch_to.frame('MainFrame')
                        self.driver.find_element(By.CSS_SELECTOR, 'input.clsImage[name="btn_ok"]').click()

                        try:
                            blank_of_user_id = self.driver.find_element(By.XPATH, '/html/body/form/div[2]/center/table/tbody/tr[1]/td/input')
                            blank_of_user_id.send_keys(user_id)
                            self.driver.find_element(By.XPATH, '/html/body/form/div[2]/center/table/tbody/tr[2]/td/input').send_keys(password)
                            self.driver.find_element(By.CSS_SELECTOR, 'input.clsImage[name="btn_ok"]').click()
                        except NoSuchElementException:
                            pass

                        try:
                            self.driver.find_element(By.NAME, "btn_next").click()
                        except Exception as e:
                            return f"{user_id}さんのコートカードは停止されている、予約制限を超えている、ID・パスワードが間違っている可能性があります。\n手動でログインして原因を確認してください。"
                        
                        self.driver.find_element(By.NAME, "btn_next").click()
                        self.driver.find_element(By.NAME, "btn_cmd").click()
                        alert = WebDriverWait(self.driver, 10).until(EC.alert_is_present())
                        alert.accept()
                        self.driver.find_element(By.NAME, "btn_back").click()
                        
                        try:
                            self.driver.find_element(By.CLASS_NAME, 'ResultMsg')
                            return f"{user_id}さんのコートカードではすでに同一コマ({lot})が予約されている、もしくはカードの有効期限が切れている可能性があります。\n手動でログインして原因を確認してください。"
                        except NoSuchElementException:
                            pass

                        logout_button = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable((By.NAME, 'btn_LogOut'))
                        )
                        logout_button.click()
                        return f"成功"
                except Exception as e:
                    return "失敗"
            else:
                return "実行がキャンセルされました。"
        else:
            return "実行がキャンセルされました。"

if __name__ == "__main__":
    url = "https://g-kyoto.pref.kyoto.lg.jp/reserve_j/core_i/init.asp?SBT=1"
    bot = CourtReserveBot(url)
    menu = st.sidebar.selectbox(label="メニュー",options=["ツール説明","入力確認・予約実行","情報入力"],key=f"menu")
    
    # 予約リストの初期化
    if "summary_list" not in st.session_state:
        st.session_state.summary_list = {
            "予約1": "",
            "予約2": "",
            "予約3": "",
            "予約4": "",
            "予約5": "",
            "予約6": "",
            "予約7": "",
            "予約8": "",
            "予約9": "",
            "予約10": ""
        }

    # フロントの実装
    if menu == "ツール説明":
        bot.landing_page()
    if menu == "入力確認・予約実行":
        bot.edit_info(st.session_state.summary_list)
        if st.button("予約を実行する"):
            for key in st.session_state.summary_list.keys():
                if st.session_state.summary_list[key] != "":
                    try:
                        st.write(f"{key}の予約を実行します。")
                        lots, park, user_id, password = st.session_state.summary_list[key]["lot"], st.session_state.summary_list[key]["park"], st.session_state.summary_list[key]["id"], st.session_state.summary_list[key]["password"]
                        success = bot.drive_website(lots=lots, park=park, user_id=user_id, password=password)
                        st.write(f"{key}の予約結果：{success}")
                    except Exception as e:
                        st.write(f"{key}の予約が失敗しました。")
    if menu == "情報入力":
        tab_labels = [f"予約{i}" for i in range(1, 11)]
        tabs = st.tabs(tab_labels)
        for i, tab in enumerate(tabs, 1):
            with tab:
                user_input_tab = bot.front(tab_name=f"tab_{i}")
                if user_input_tab != None:
                    st.session_state.summary_list[f"予約{i}"] = user_input_tab