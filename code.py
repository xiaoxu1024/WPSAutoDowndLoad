from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
import time
import os
import xlwt,xlrd

# 定义一些需要的数据
# wps_username=17334033685                     # wps的账号
# wps_password='Wei17334033685'                     # wps的密码
url = 'https://www.docer.com/login'      # 运行的网址

FILE = "././"        # 储存路径

def loginAndsearch(url,wps_username,wps_password):
    # 1.创建、配置并启动chrome

    # 创建一个chrome
    options = webdriver.ChromeOptions()

    # 配置浏览器
    prefs = {
        'profile.managed_default_content_settings.images': 2,            # 不加载浏览器的图片
        "credentials_enable_service": False,                             # 浏览器弹窗
        "profile.password_manager_enabled": False,                       # 关闭浏览器弹窗
        "download.default_directory": "e:\\WPSGET"
    }
    options.add_experimental_option('prefs', prefs)                     # 参数送入执行
    # 设置为开发者模式，防止被各大网站识别出来使用了Selenium
    options.add_experimental_option('excludeSwitches', ['enable-automation'])

    # 启动这个chrome
    browser = webdriver.Chrome(options=options)
    wait = WebDriverWait(browser, 10)                                   # 超时时长为10s

    # 2.打开网页
    browser.get(url)

    # 转到iframe
    browser.implicitly_wait(30)
    elementi = browser.find_element_by_class_name('ifm')
    browser.switch_to.frame(elementi)

    # 选择账号密码输入
    account_click='//span[contains(text(),"帐号密码")]'
    browser.find_element_by_xpath(account_click).click()

    # 点击同意协议
    browser.find_element_by_xpath('//div[@class="dialog-footer-ok"]').click()

    # 键入账号和密码
    browser.find_element_by_xpath('//input[@id="email"]').send_keys(wps_username)       # 键入账号
    browser.find_element_by_xpath('//input[@id="password"]').send_keys(wps_password)    # 键入密码

    # 点击验证
    browser.implicitly_wait(30)
    browser.find_element_by_xpath('//div[@id="rectMask"]').click()

    # 点击登录
    time.sleep(5)
    browser.find_element_by_xpath('//a[@id="login"]').click()

    browser.switch_to.default_content()

    # 点击回到首页
    time.sleep(2)
    browser.implicitly_wait(30)
    browser.find_element_by_xpath('//a[@class="nav_li_a "]').click()

    return browser

# 创建新的文件夹
def creatFile(element, FILE=FILE):
    path = FILE
    title = element
    new_path = os.path.join(path, title)
    if not os.path.isdir(new_path):
        os.makedirs(new_path)
    return new_path

def creatExcel(model_label):
    workbook = xlwt.Workbook(encoding='utf-8')  # 新建工作簿
    sheet1 = workbook.add_sheet(model_label)  # 新建sheet
    return workbook,sheet1



# PPT页面操作
def PPTDownload(browser,model_label):
    print("ppt")

    # 创建关键词构成的文件夹
    creatFile(model_label, FILE=FILE+'ppt/')
    # 创建关键词构成的excel
    workbook,sheet1=creatExcel(model_label)

    # 打开ppt模板页
    browser.implicitly_wait(30)
    browser.find_element_by_xpath('//*[@id="App"]/div[2]/div[2]/ul/li[1]/ul/li[3]/a').click()
    time.sleep(3)               # 等待页面加载完成

    page = browser.find_element_by_xpath('//*[@id="App"]/div[2]/div[4]/div[2]/span').text
    print('总页数：', page[1:])

    count=0
    for p in range(int(page[1:])):
        num=browser.find_element_by_xpath('//*[@id="App"]/div[2]/div[4]').get_attribute('len')
        print('page'+str(p+1)+':',num,'个')
        time.sleep(2)

        for i in range(1, int(num) + 1):
            browser.implicitly_wait(30)
            href = browser.find_element_by_xpath(
                '//*[@id="App"]/div[2]/div[4]/ul' + '/li[' + str(i) + ']/a').get_attribute('href')

            browser.implicitly_wait(30)
            title = browser.find_element_by_xpath(
                '//*[@id="App"]/div[2]/div[4]/ul' + '/li[' + str(i) + ']/a/div[2]').get_attribute('title')

            # 写入数据
            sheet1.write(count, 0, title)  # 第1行第1列数据
            sheet1.write(count, 1, href)  # 第1行第2列数据
            count=count+1
            print(title, ' ', href)
        workbook.save(r'././' + 'ppt' + '/' + model_label + '/' + model_label + '.xls')
        browser.implicitly_wait(30)
        browser.find_element_by_xpath('//*[@id="App"]/div[2]/div[4]/div[2]/a[4]').click()
        time.sleep(2)

    return browser

# word页面操作
def WORDDownload(browser,model_label):
    print('word')

    # 创建关键词构成的文件夹
    creatFile(model_label, FILE=FILE + 'word/')
    # 创建关键词构成的excel
    workbook, sheet1 = creatExcel(model_label)

    # 打开ppt模板页
    browser.implicitly_wait(30)
    browser.find_element_by_xpath('//*[@id="App"]/div[2]/div[2]/ul/li[1]/ul/li[2]/a').click()
    time.sleep(3)  # 等待页面加载完成

    page = browser.find_element_by_xpath('//*[@id="App"]/div[2]/div[4]/div[2]/span').text
    print('总页数：', page[1:])

    count = 0
    for p in range(int(page[1:])):
        num = browser.find_element_by_xpath('//*[@id="App"]/div[2]/div[4]').get_attribute('len')
        print('page' + str(p + 1) + ':', num, '个')
        time.sleep(2)

        for i in range(1, int(num) + 1):
            browser.implicitly_wait(30)
            href = browser.find_element_by_xpath(
                '//*[@id="App"]/div[2]/div[4]/ul' + '/li[' + str(i) + ']/a').get_attribute('href')

            browser.implicitly_wait(30)
            title = browser.find_element_by_xpath(
                '//*[@id="App"]/div[2]/div[4]/ul' + '/li[' + str(i) + ']/a/div[2]').get_attribute('title')

            # 写入数据
            sheet1.write(count, 0, title)  # 第1行第1列数据
            sheet1.write(count, 1, href)  # 第1行第2列数据
            count = count + 1
            print(title, ' ', href)
        workbook.save(r'././' + 'word' + '/' + model_label + '/' + model_label + '.xls')

        browser.implicitly_wait(30)
        browser.find_element_by_xpath('//*[@id="App"]/div[2]/div[4]/div[2]/a[4]').click()
        time.sleep(2)

    return browser



# excel页面操作
def EXCELDownload(browser,model_label):
    print('excel')

    # 创建关键词构成的文件夹
    creatFile(model_label, FILE=FILE + 'excel/')
    # 创建关键词构成的excel
    workbook, sheet1 = creatExcel(model_label)

    # 打开ppt模板页
    browser.implicitly_wait(30)
    browser.find_element_by_xpath('//*[@id="App"]/div[2]/div[2]/ul/li[1]/ul/li[4]/a').click()
    time.sleep(3)  # 等待页面加载完成

    page = browser.find_element_by_xpath('//*[@id="App"]/div[2]/div[4]/div[2]/span').text
    print('总页数：', page[1:])

    count = 0
    for p in range(int(page[1:])):
        num = browser.find_element_by_xpath('//*[@id="App"]/div[2]/div[4]').get_attribute('len')
        print('page' + str(p + 1) + ':', num, '个')
        time.sleep(2)

        for i in range(1, int(num) + 1):
            browser.implicitly_wait(30)
            href = browser.find_element_by_xpath(
                '//*[@id="App"]/div[2]/div[4]/ul' + '/li[' + str(i) + ']/a').get_attribute('href')

            browser.implicitly_wait(30)
            title = browser.find_element_by_xpath(
                '//*[@id="App"]/div[2]/div[4]/ul' + '/li[' + str(i) + ']/a/div[2]').get_attribute('title')

            # 写入数据
            sheet1.write(count, 0, title)  # 第1行第1列数据
            sheet1.write(count, 1, href)  # 第1行第2列数据
            count = count + 1
            print(title, ' ', href)
        workbook.save(r'././' + 'excel' + '/' + model_label + '/' + model_label + '.xls')

        browser.implicitly_wait(30)
        browser.find_element_by_xpath('//*[@id="App"]/div[2]/div[4]/div[2]/a[4]').click()
        time.sleep(2)

    return browser








# 文件名和路径转换
def oldTonew_PPT(name,flage_num):
    file_path = "e:/WPSGET/"
    file = os.listdir(file_path)
    # 更改名字和路径
    time.sleep(0.5)
    for f in range(len(file)):
        type = os.path.splitext(file[f])[1]
        if (type == '.pptx' or type == '.ppt' or type=='.dpt'):
            print(file[f])
            # 获取旧文件名
            oldname = file_path + file[f]  # os.sep添加系统分隔符
            # 设置新文件名
            newname = 'e:/WPSGET/ppt/'+model_label+'/'+str(flage_num)+ name
            os.rename(oldname, newname)  # 用os模块中的rename方法对文件改名
            print('下載完成')
def oldTonew_WORD(name,flage_num):
    file_path = "e:/WPSGET/"
    file = os.listdir(file_path)
    # 更改名字和路径
    time.sleep(0.5)
    for f in range(len(file)):
        type = os.path.splitext(file[f])[1]
        if (type == '.docx' or type == '.doc' or type=='.wpt' or type=='wps'):
            print(file[f])
            # 获取旧文件名
            oldname = file_path + file[f]  # os.sep添加系统分隔符
            # 设置新文件名
            newname = 'e:/WPSGET/word/'+model_label+'/'+ str(flage_num)+name
            os.rename(oldname, newname)  # 用os模块中的rename方法对文件改名
            print('下載完成')

def oldTonew_EXCEL(name,flage_num):
    file_path = "e:/WPSGET/"
    file = os.listdir(file_path)
    # 更改名字和路径
    time.sleep(0.5)
    for f in range(len(file)):
        type = os.path.splitext(file[f])[1]
        if (type == '.xlsx' or type == '.xls' or type=='.csv' or type=='ett'):
            print(file[f])
            # 获取旧文件名
            oldname = file_path + file[f]  # os.sep添加系统分隔符
            # 设置新文件名
            newname = 'e:/WPSGET/excel/'+model_label+'/'+ str(flage_num)+name
            os.rename(oldname, newname)  # 用os模块中的rename方法对文件改名
            print('下載完成')
# 下载
def Download(browser,type,model_label):

    excle_path = './'+type+'/'+model_label+'/'+model_label+'.xls'           # excel路径
    data = xlrd.open_workbook(excle_path)       # 打开excel读取文件
    sheet = data.sheet_by_index(0)              # 根据sheet下标选择读取内容
    nrows = sheet.nrows                         # 获取到表的总行数
    print(nrows)
    for j in range(int(nrows)):
        name=sheet.row_values(j)[0]
        ll = sheet.row_values(j)[1]

        browser.get(ll)
        time.sleep(1)
        browser.implicitly_wait(40)
        # browser.find_element_by_xpath('//*[@id="dlBtn"]').click()

        ele = browser.find_element_by_xpath('//*[@id="dlBtn"]')
        browser.execute_script("arguments[0].click();", ele)

        if(type=='ppt'):
            time.sleep(15)
            oldTonew_PPT(name,j)
        elif(type=='word'):
            time.sleep(10)
            oldTonew_WORD(name,j)
        elif(type=='excel'):
            time.sleep(10)
            oldTonew_EXCEL(name, j)



    return 1

###################################################################

if __name__ == "__main__":

    # 提示输入账号和密码
    wps_username = input("请输入账号：")
    wps_password = input("请输入密码：")

    # 登录wps
    browser=loginAndsearch(url,wps_username,wps_password)
    # 输入关键词
    model_label = input("请输入关键词：")
    browser.find_element_by_xpath('//div[@class="m-search-box header-banner__search"]/input').send_keys(
        model_label + '\n')
    time.sleep(1)
    # 切换页面后必须切换句柄，不然找不到元素
    num =browser.window_handles         # 获取当前页句柄
    browser.switch_to.window(num[1])    # 在句柄2 上执行下述步骤

    # 分类分页下载
    model_type = input("请输入需要下载的模板类型：")

    if(model_type=='ppt' or model_type=='PPT'):
        browser=PPTDownload(browser,model_label)
        flage=Download(browser,'ppt',model_label)

    elif(model_type=='word' or model_type=='WORD'):
        browser=WORDDownload(browser,model_label)
        flage = Download(browser,'word',model_label)
    elif(model_type=='excel' or model_type=='EXCEL'):
        browser=EXCELDownload(browser,model_label)
        flage = Download(browser, 'excel', model_label)
    else:
        print('输入类型错误')




