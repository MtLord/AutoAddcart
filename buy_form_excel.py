from argparse import ArgumentParser as Parser
import openpyxl

from selenium import webdriver
import chromedriver_binary
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
url = []#URL
buy_quan = []#購入数量
site = []#購入サイト名
prduct_name = []#製品名
driver = webdriver.Chrome()# chromeを開く

def ListXlsElemet(url_row, buy_quan_row, site_row, proname_row): 
    parser = Parser(description='EXCELに書かれた商品リストを自動的にカートに入れるスクリプト')
    parser.add_argument("path", type=str, help='excelファイルのパスを指定')
    args = parser.parse_args()
    wb = openpyxl.load_workbook(args.path)
    sheet = wb['Sheet1']
    for sell in sheet.iter_rows(min_row=2, max_col=8):
        values = []
        for col in sell:
            values.append(col.value)
        if values[url_row] != None:
            url.append(values[url_row])
        if values[buy_quan_row] != None:
            buy_quan.append(values[buy_quan_row])
        if values[site_row] != None:
            site.append(values[site_row])
        prduct_name.append(values[proname_row])


class GoToCart:
    def __init__(self, arg_url, arg_quan, arg_site,arg_proname):
   
        self.url = arg_url
        self.buy_quan = arg_quan
        self.site = arg_site
        self.proname=arg_proname
    def CharClassifi(self, ch, target):  # 特定の文字列を含むかどうかを判定し、一致したらリストに追加する関数
        relist = []
        num = 0
        for char in self.site:
            if str(char) == ch:
                relist.append(target[num])

            num = num + 1
        return relist

    # サイト名とxpthsを渡すと必要数分カートに入れる関数
    def AddCart(self, site_name, num_xpath, cart_xpath):
        temp_url = self.CharClassifi(site_name, self.url)
        temp_quan = self.CharClassifi(site_name, self.buy_quan)
        temp_proname = self.CharClassifi(site_name, self.proname)
        num = 0
        helth=0
        for t_url in temp_url:
            try:
                driver.get(t_url)
                wait = WebDriverWait(driver, 10)
                # 購入数量のxpathを指定する
                t_element = driver.find_element_by_xpath(num_xpath)
                t_element.clear()
                t_element.send_keys(str(temp_quan[num]))
                c_element = driver.find_element_by_xpath(cart_xpath)
                c_element.click()
                
                
            except:
                    print(temp_proname[num])
                    print("カートに入れられません。商品が存在しないか在庫がない可能性があります")
                    continue
            num = num+1

class ChildGotoCart(GoToCart):
    def Akiduki(self):
        self.AddCart( "秋月", '//*[@id="maincontents"]/div[2]/table/tbody/tr/td[3]/input','//*[@id="maincontents"]/div[2]/table/tbody/tr/td[4]/input')

    def Sengoku(self):
        driver.execute_script("window.open()")
        driver.switch_to.window(driver.window_handles[1])#新規タブに切り替え
        self.AddCart("千石",'//*[@id="detail_price"]/tbody/tr/td[2]/input','//*[@id="detail_price"]/tbody/tr/td[3]/input')
    





def main():
    buy_handle = ChildGotoCart(url, buy_quan, site,prduct_name)
    ListXlsElemet(7, 6, 3,2)
    buy_handle.Akiduki()
    buy_handle.Sengoku()
    print("完了しました")
    return 0


if __name__ == "__main__":
    main()
