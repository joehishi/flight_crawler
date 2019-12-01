from pudong_utils import PD_spider
from daxing_utils import DX_spider
import argparse
import platform
import sys

firefox_linux = './geckodriver'
firefox_win = './geckodriver-v0.26.0-win64/geckodriver.exe'
firefox_mac = '/Users/huaiyuchen/Downloads/selenium_rel/geckodriver'

'''
def parse_args():
    parser = argparse.ArgumentParser(description='Fligtht Crawler')
    parser.add_argument('-t', '--type',  help='input PD or DX')
    parser.add_argument('-d', '--day',  help='True for today otherwise False')

    args = parser.parse_args()
    return args
'''
if(platform.system() == 'Linux'):
    driver_path = firefox_linux
if(platform.system() == 'Darwin'):
    driver_path = firefox_mac
else:
    driver_path = firefox_win

#args = parse_args()

if(sys.argv[1] == 'PD'):
    spider = PD_spider(driver_path = driver_path,Today = sys.argv[2])
    spider.run_all()

if(sys.argv[1] == 'DX'):
    spider = DX_spider(driver_path = driver_path,Today = sys.argv[2])
    spider.run_all()
