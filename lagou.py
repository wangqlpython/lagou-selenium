from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from lxml import etree
import time
import re
import pandas as pd
import xlwings as xw


class LaGou():
    def __init__(self):
        self.opt = Options()
        self.opt.add_argument("start-maximized")
        self.driver=webdriver.Chrome(options=self.opt)
        self.url = "https://www.lagou.com/jobs/list_%20python?labelWords=&fromSearch=true&suginput="
        self.positions = []
        self.count_num = 1
        self.save_count = 0
        self.app = xw.App(visible=True,add_book=False)
        self.position_file = self.app.books.open("position.xlsx")
        self.sheet= self.position_file.sheets[0]

    def run(self):
        #打开网页
        self.driver.get(self.url)
        #登录
        self.login()
        #从多少页开始爬取
        spider_page = int(input("输入从第几页开始爬取，输入整数："))
        if spider_page>1:
            self.continue_spider(spider_page)
        while True:
            wait(self.driver,timeout=10).until(EC.presence_of_all_elements_located((By.XPATH,"//div[@class='pager_container']/span[last()]")))
            source = self.driver.page_source
            # print(source)
            self.parse_page_url(source)

            next_page = self.driver.find_element_by_xpath("//div[@class='pager_container']/span[@action='next']")
            adv_page= self.driver.find_element_by_xpath("//div[@class='body-btn']")

            if adv_page.text:
                adv_page.click()

            if 'pager_next pager_next_disabled' in next_page.get_attribute("class"):
                print("爬取完成")
                break
            else:
                next_page.click()
                time.sleep(1)


    def parse_page_url(self,source):
        html = etree.HTML(source)
        detail_links = html.xpath("//a[@class='position_link']/@href")
        print(detail_links)
        for link in detail_links:
            self.request_detail_page(link)
            time.sleep(1)
            # break

    def request_detail_page(self,url):
        self.driver.execute_script("window.open('%s')" % url)
        self.driver.switch_to.window(self.driver.window_handles[1])
        wait(self.driver,timeout=10).until(EC.presence_of_all_elements_located((By.XPATH,"//span[@class='name']")))
        page_source = self.driver.page_source
        self.parse_detail_page(page_source)
        time.sleep(1)
        self.driver.close()
        self.driver.switch_to.window(self.driver.window_handles[0])

    def parse_detail_page(self,source):
        html = etree.HTML(source)
        company_name = html.xpath("//h4[@class='company']/text()")[0]
        position_name = html.xpath("//div[@class='job-name']/@title")[0]
        job_request = html.xpath("//dd[@class='job_request']//span")
        # print(job_request)
        salary = job_request[0].xpath(".//text()")[0].strip()
        city = job_request[1].xpath(".//text()")[0].strip()
        city = re.sub("[\s/]","",city)
        experience = job_request[2].xpath(".//text()")[0].strip()
        experience=re.sub("[\s/]","",experience)
        education = job_request[3].xpath(".//text()")[0].strip()
        education = re.sub("[\s/]", "", education)
        full_or_part=job_request[4].xpath(".//text()")[0].strip()
        full_or_part = re.sub("[\s/]", "", full_or_part)

        job_advantage = html.xpath("//dd[@class='job-advantage']/p/text()")[0].strip()
        job_describe = html.xpath("//dd[@class='job_bt']//text()")
        job_describe = [re.sub("[\s]","",x) for x in job_describe if len(re.sub("[\s]","",x))>0]

        job_describe="".join(job_describe)
        # print(job_describe)
        position = {
            'company_name': company_name,
            'position_name': position_name,
            'salary': salary,
            'city': city,
            'experience': experience,
            'education': education,
            'full_or_part': full_or_part,
            'job_describe': job_describe,
            'job_advantage': job_advantage
        }
        print(position)
        self.positions.append(position)
        self.count_num += 1
        if self.count_num%15==1:
            self.save_positions()



    def save_positions(self):
        save_positions=pd.DataFrame(self.positions)
        self.positions=[]
        row = 1 + 16*self.save_count
        self.save_count += 1
        print("已保存%d页" %self.save_count)
        print("*"*30)
        if self.save_count==1:
            self.sheet.range("A"+str(row)).value=save_positions
        else:
            self.sheet.range("B" + str(row)).value = save_positions.values
        self.position_file.save()

    def login(self):
        adv_page = self.driver.find_element_by_xpath("//div[@class='body-btn']")
        adv_page.click()
        loginTag = self.driver.find_element_by_css_selector(".login")
        usernameTag = self.driver.find_element_by_xpath("//input[@type='text']")
        passwordTag = self.driver.find_element_by_xpath("//input[@type='password']")
        login = self.driver.find_element_by_xpath(
            "//div[@class='login-btn login-password sense_login_password btn-green']")

        actions = ActionChains(self.driver)
        actions.move_to_element(loginTag)
        actions.click(loginTag)
        actions.send_keys_to_element(usernameTag,"18908228467")
        actions.send_keys_to_element(passwordTag,"wql513624")
        actions.move_to_element(login)
        actions.click(login)
        actions.perform()
        time.sleep(15)

    def continue_spider(self,num):
        self.count_num = 15*num+1
        self.save_count = num-1

        current_page = 1
        while True:
            if current_page==num:
                break
            else:
                next_page_Btn = self.driver.find_element_by_xpath("//div[@class='pager_container']/span[last()]")
                actions = ActionChains(self.driver)
                actions.move_to_element(next_page_Btn)
                actions.click(next_page_Btn)
                actions.perform()

                current_page += 1
                time.sleep(1)



def main():
    lagou=LaGou()
    lagou.run()

if __name__=="__main__":
    main()