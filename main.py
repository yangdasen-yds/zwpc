"""
-------------------------------------------------
   File Name：     main.py
   date：          2020/11/23
-------------------------------------------------
   通过高级检索爬取知网论文摘要(使用的是旧版网址:https://kns.cnki.net/kns/brief/result.aspx）
                   
-------------------------------------------------
"""
import os
import re
import shutil
import time
# 引入字节编码
from urllib.parse import quote
import xlwt
import requests
import urllib3
# 引入beautifulsoup
from bs4 import BeautifulSoup

# 解决访问网站时的警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

HEADER = {
    'User-Agent':
        'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Safari/537.36',
}
# 获取cookie
# BASIC_URL = 'https://kns.cnki.net/kns/brief/result.aspx'
BASIC_URL = 'https://kns.cnki.net/kns8/AdvSearch?'
# 利用post请求先行注册一次
SEARCH_HANDLE_URL = 'https://kns.cnki.net/kns/request/SearchHandler.ashx'
# 发送get请求获得文献资源
GET_PAGE_URL = 'https://kns.cnki.net/kns/brief/brief.aspx?pagename='
# DOWNLOAD_URL = 'https://kns.cnki.net/kns/'
# 切换页面基础链接
CHANGE_PAGE_URL = 'https://kns.cnki.net/kns/brief/brief.aspx'
# 论文信息基础链接
aa1 = 'https://kns.cnki.net/KCMS/detail/detail.aspx?'
# 设置检索参数
userdata = {'txt_1_sel': 'SU$%=|',
            'txt_1_value1': '数学建模',
            'txt_1_relation': '#CNKI_AND',
            'txt_1_special1': '=',  # (=:表示精确)
            'txt_2_sel': 'TI',
            'txt_2_value1': '数学建模',
            'txt_2_logical': 'or',
            'txt_2_relation': '#CNKI_AND',
            'txt_2_special1': '=',
            'txt_3_sel': 'KY',
            'txt_3_value1': '数学建模',
            'txt_3_logical': 'or',
            'txt_3_relation': '#CNKI_AND',
            'txt_3_special1': '%',  # (%:表示模糊)
            'txt_4_sel': 'FT',
            'txt_4_value1': '中学',
            'txt_4_logical': 'and',
            'txt_4_relation': '#CNKI_AND',
            'txt_4_special1': '='
            }


class SearchTools(object):
    """

    构建搜索类
    实现搜索方法

    """

    def __init__(self):
        # 获取文献摘要等信息存储至excel
        self.xls = xlwt.Workbook(encoding='utf-8')
        self.sheet = self.xls.add_sheet("shttl1", cell_overwrite_ok=True)
        self.sheet.write(0, 0, '篇名')
        self.sheet.write(0, 1, '摘要')
        self.number = 1
        ###
        self.session = requests.Session()
        self.cur_page_num = 1
        # 保持会话
        self.session.get(BASIC_URL, headers=HEADER)

    def search_reference(self, ueser_input):
        # 检索参数
        static_post_data = {
            'action': '',
            'NaviCode': '*',
            'ua': '1.21',
            'isinEn': '1',
            'PageName': 'ASP.brief_default_result_aspx',
            'DbPrefix': 'SCDB',
            'DbCatalog': '中国学术期刊网络出版总库',
            'ConfigFile': 'CJFQ.xml',
            'db_opt': 'CJFQ,CDFD,CMFD,CPFD,IPFD,CCND,CCJD',  # 搜索类别（CNKI右侧的）
            'db_value': '中国学术期刊网络出版总库',
            '@joursource': '( 核心期刊=Y or CSSCI期刊=Y)',  # (选择核心期刊和CSSCI期刊)
            'year_type': 'echar',
            'his': '0',
            'db_cjfqview': '中国学术期刊网络出版总库,WWJD',
            'db_cflqview': '中国学术期刊网络出版总库',
            '__': time.asctime(time.localtime()) + ' GMT+0800 (中国标准时间)'
        }
        # 拼接static_post_data, ueser_input
        post_data = {**static_post_data, **ueser_input}
        # 必须有第一次请求，否则会提示服务器没有用户
        first_post_res = self.session.post(
            SEARCH_HANDLE_URL, data=post_data, headers=HEADER)
        print(first_post_res.text)
        # get请求中需要传入第一个检索条件的值
        key_value = quote(ueser_input.get('txt_1_value1'))
        self.get_result_url = GET_PAGE_URL + first_post_res.text + '&t=1544249384932&keyValue=' + key_value + '&S=1&sorttype='
        # 检索结果的第一个页面
        print(self.get_result_url)
        second_get_res = self.session.get(self.get_result_url, headers=HEADER)
        # 翻页URL准备
        change_page_pattern_compile = re.compile(
            r'.*?pagerTitleCell.*?<a href="(.*?)".*')

        try:
            self.change_page_url = re.search(change_page_pattern_compile,
                                             second_get_res.text).group(1)
        except:
            pass
        ###
        self.parse_page(
            self.pre_parse_page(second_get_res.text), second_get_res.text)
        # 保存
        self.xls.save(r'data/title+abstract.xls')

    def pre_parse_page(self, page_source):
        """
        选择需要检索的页数
        """
        reference_num_pattern_compile = re.compile(r'.*?找到&nbsp;(.*?)&nbsp;')
        reference_num = re.search(reference_num_pattern_compile,
                                  page_source).group(1)
        reference_num_int = int(reference_num.replace(',', ''))
        print('检索到' + reference_num + '条结果，全部下载大约需要' +
              s2h(reference_num_int * 5) + '。')
        is_all_download = input('是否要全部下载（y/n）?')
        # 将所有数量根据每页20计算多少页
        if is_all_download == 'y':
            page, i = divmod(reference_num_int, 20)
            if i != 0:
                page += 1
            return page
        else:
            select_download_num = int(input('请输入需要下载的数量（不满一页将下载整页）：'))
            while True:
                if select_download_num > reference_num_int:
                    print('输入数量大于检索结果，请重新输入！')
                    select_download_num = int(input('请输入需要下载的数量（不满一页将下载整页）：'))
                else:
                    page, i = divmod(select_download_num, 20)
                    # 不满一页的下载一整页
                    if i != 0:
                        page += 1
                    print("开始下载前%d页所有文件，预计用时%s" % (page, s2h(page * 20 * 5)))
                    print('－－－－－－－－－－－－－－－－－－－－－－－－－－')
                    return page

    def parse_page(self, download_page_left, page_source):
        '''
        解析论文信息URL
        '''
        print('正在请求第----------------' + str(self.cur_page_num) + '页')
        soup = BeautifulSoup(page_source, 'lxml')
        list = soup.find_all(class_='fz14')
        for li in list:
            # 正则解析
            li = str(li)
            # print(li)
            reference_num_pattern_compile = re.compile('recid=&amp;(.*?)&amp;yx=')
            reference_num = re.search(reference_num_pattern_compile,
                                      li).group(1)
            url1 = reference_num.replace('amp;', '')
            url = aa1 + url1
            print('论文信息地址:' + url)
            self.download(url)

        if download_page_left > 1:
            # print(download_page_left)
            self.cur_page_num += 1
            self.get_another_page(download_page_left)

    def get_another_page(self, download_page_left):
        '''
        请求其他页面和请求第一个页面形式不同
        重新构造请求
        '''
        curpage_pattern_compile = re.compile(r'.*?curpage=(\d+).*?')
        self.get_result_url = CHANGE_PAGE_URL + re.sub(
            curpage_pattern_compile, '?curpage=' + str(self.cur_page_num),
            self.change_page_url)

        self.get_res = self.session.get(self.get_result_url, verify=False, headers=HEADER)
        # print(self.get_result_url)
        download_page_left -= 1
        self.parse_page(download_page_left, self.get_res.text)

    def download(self, detailurl):
        """

        下载论文篇名和摘要，保存至data/abstract.txt

        """
        self.response = self.session.get(detailurl, verify=False, headers=HEADER).text
        print('请求成功............')
        soupone = BeautifulSoup(self.response, 'lxml')
        title = soupone.find('div', class_='wx-tit').h1.text
        title = str(title)
        # 篇名写入xls
        self.sheet.write(self.number, 0, title)
        try:
            abstract = soupone.find('span', class_='abstract-text').text
        except:
            abstract = '无摘要'
        with open('data/abstract.txt', 'a', encoding='utf-8') as file:
            file.write(abstract + '\n\n\n')
        abstract = str(abstract)
        # 摘要写入xls
        self.sheet.write(self.number, 1, abstract)
        self.number += 1

        print('篇名:' + title)
        # print(abstract)


def s2h(seconds):
    '''
    将秒数转为小时数
    '''
    m, s = divmod(seconds, 60)
    h, m = divmod(m, 60)
    return "%02d小时%02d分钟%02d秒" % (h, m, s)


def main():
    time.perf_counter()
    if os.path.isdir('data'):
        # 递归删除文件
        shutil.rmtree('data')
    # 创建一个空的
    os.mkdir('data')
    search = SearchTools()
    search.search_reference(userdata)
    print('－－－－－－－－－－－－－－－－－－－－－－－－－－')
    print('爬取完毕，共运行：' + s2h(time.perf_counter()))


if __name__ == '__main__':
    main()
