import xlwings as xw
import config

wb = xw.Book(repair_excel)

sht = wb.sheets['作業記録']
sht.activate() #作業記録シートを開き、アクティブシートにする

stf_name = sht.range('C694').value
#.replace('　',' ')C694の氏名の全角スペースを半角に変換し、nameに格納

driver = webdriver.Chrome(executable_path=ChromeDriverManager().install()) #Chromeを開く
driver.get(spfa_top) #SPFAを開く
time.sleep(3)

user_name = driver.find_element(By.NAME,'USER_NAME') #ログインID入力欄に移動
user_name.send_keys(spfa_id) #ログインIDにのID入力

password = driver.find_element(By.NAME,'PASSWORD') #パスワード入力欄に移動
password.send_keys(spfa_pass) #パスワード入力 中身後で環境変数にする
driver.find_element(By.XPATH,'/html/body/form/table/tbody/tr[6]/td/input').click() #ログインボタンを押す
time.sleep(3)

driver.get(res_calc) #リンクを踏ませると何故か正常に動かないので直接遷移
time.sleep(3)

f_date = driver.find_element(By.CSS_SELECTOR,"input[type='TEXT'][name='S_DATE']")
f_date.send_keys('3/25') #検索初めの日に本日の日付入力

l_date = driver.find_element(By.CSS_SELECTOR,"input[type='TEXT'][name='L_DATE']")
l_date.send_keys('3/25') #検索終わりの日に本日の日付入力

driver.find_element(By.CSS_SELECTOR,"input[type=button][name='SUB1']").click() #検索開始のボタンを押す
time.sleep(5)

driver.find_element(By.LINK_TEXT, stf_name).click()
time.sleep(5)



xl = xw.apps.active.api
xl.Quit()
