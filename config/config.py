from bs4 import BeautifulSoup
from pathlib import Path

config_folder = str(Path().absolute()) + '\\config'
config_file = 'Config.xml'

class Config:

    def __init__(self):
        bs = self.initBS() 
        self.init(bs)

    def init(self, bs):
        self.email_user     = bs.find('email').find('user').text
        self.email_pass     = bs.find('email').find('password').text
        self.email_server   = bs.find('email').find('server').text
        self.inter_url      = bs.find('inter').find('url').text
        self.inter_user     = bs.find('inter').find('user').text
        self.inter_pass     = bs.find('inter').find('password').text
        self.inter_sent     = bs.find('inter').find('sent').text
        self.inter_extract  = bs.find('inter').find('extract-file').text
        self.path_download  = bs.find('path').find('downloads').text
        self.path_prd       = bs.find('path').find('prd').text
        self.path_backup    = bs.find('path').find('backup').text

    def initBS(self):
        with open(config_folder + '\\' + config_file) as fp:
            soup = BeautifulSoup(fp, 'lxml')

        return soup
