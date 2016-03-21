
from requests import get
from bs4 import BeautifulSoup
from os import listdir, mkdir, chdir, path
from wget import download
from zipfile import ZipFile
from urllib import ContentTooShortError
from subprocess import Popen

class GetShapeFiles:
    def __init__(self):
#        self.dirname = 'C:\Users\User\Downloads'
        self.page_url = 'http://www.gadm.org/country'
        self.file_names = []
        self.current_file_dir = path.dirname(path.abspath(__file__))
        self.dirname = self.current_file_dir
 
    def build_dest(self, dirname):
        mkdir(dirname)
        mkdir(dirname + '/zipfiles')
        return self.dirname + '/' + dirname

    def gadm_download_files(self):
        data = get(self.page_url)
        soup = BeautifulSoup(data.text, 'html.parser')
        dirpath = self.build_dest('gadmzips')
        chdir(dirpath + '/zipfiles')
        select_childs = soup.find('select').findChildren()
        for i in select_childs:
            try:
                raw_file_name = i['value'].split('_')[0] + '_adm_shp'
                zip_file_name = i['value'].split('_')[0] + '_adm_shp.zip'
                self.file_names.append(raw_file_name)
                download('http://biogeo.ucdavis.edu/data/gadm2.8/shp/' + zip_file_name)

            except ContentTooShortError as e:
                print (e.__str__())
            except Exception as e:
                print (e.__str__())
        chdir(self.dirname)
        return dirpath

    def unzip_gadm_files(self, dirpath):
        for fl in self.file_names:
            with open(dirpath + '/zipfiles/' + fl + '.zip'):
                zippedFiles = ZipFile(dirpath + '/zipfiles/' + fl + '.zip')
                zippedFiles.extractall(dirpath + '/' + fl)

    def create_topojson(self, dirpath):
        list_of_dirs = [i for i in listdir(dirpath) if i != 'zipfiles']
        count = 0
        for dr in list_of_dirs:
            shp_dir = dirpath+'/'+dr
            for i in self.topojson_generator(shp_dir, count):
                count = i

    def topojson_generator(self, shp_dir, count):
        file_str = ''
        chdir(shp_dir)
        for f in listdir(shp_dir):
            if f.endswith('.shp'):
                file_str += f + ' '
        file_str = file_str[:len(file_str) - 1]
        if file_str:
            cmd = 'topojson' + ' ' + file_str + ' ' + '  -o ' + path.basename(shp_dir) + '.json'
            process = Popen(cmd, shell=True)
            try:
                process.wait()
                chdir(self.dirname)
                cmd = 'python shape.py ' + shp_dir + '/' +path.basename(shp_dir) +'.json'
                print (cmd)
                process = Popen(cmd, shell=True)
                process.wait()
            except:
                print ('Timout for executing the process.')
                process.kill()
            finally:
                chdir(dirpath)
        yield count + 1

if __name__ == '__main__':
    gsf = GetShapeFiles()
    dirpath = gsf.gadm_download_files()
    gsf.unzip_gadm_files(dirpath)
    gsf.create_topojson(dirpath)