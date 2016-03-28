
try:
    from requests import get
    from lxml import html
    from os import listdir, makedirs, chdir, path
    from urllib import urlretrieve
    from zipfile import ZipFile
    from urllib import ContentTooShortError
    from subprocess import Popen
    import shlex

except ImportError as e:
    print (e.__str__())
    exit()


class GetShapeFiles:
    '''
        Class GetShapeFiles is used to generate topojson from
        shapefiles and then able to pass the topojson to
        shape.py to create excel maps. It downloads shapefiles
        from gadm resource and parse them into topojson format.
    '''

    def __init__(self):
        '''
            Function to initialize class object and store object
            variables which can be used any time for other methods.
        '''

        self.page_url = 'http://www.gadm.org/country'
        self.file_names = []
        self.dirname = path.dirname(path.abspath(__file__))

    def build_dest(self, dirname):
        '''
            Function to create directory structure to store the
            downloaded files into the project directory. It takes
            dirname as an argument name which is parent directory
            name in which all the shape files get stored.
        '''

        if not path.exists(dirname + '/zipfiles'):
            makedirs(dirname + '/zipfiles')

        return self.dirname + '/' + dirname

    def gadm_download_files(self):
        '''
            Function is used to download the shape files from gadm
            resource. It calls 'build_dest()' internally to create
            directory structure and then start downloading the shape
            files. It stores all zipfiles into 'zipfiles' folder.

            It calls 'unzip_gadm_files()' for unzipping the downlaoded
            shape files zip folder and store the unzipped directories one
            step up from the zip folder.
        '''

        data = get(self.page_url)
        tree = html.fromstring(data.content)
        dirpath = self.build_dest('gadmzips')
        select_childs = tree.xpath('//select[@name="cnt"]/option')
        for i in select_childs[:1]:
            try:
                raw_file_name = i.values()[0].split('_')[0] + '_adm_shp'
                zip_file_name = raw_file_name + '.zip'
                self.file_names.append(raw_file_name)
                if path.isdir(zip_file_name):
                    pass
                else:
                    urlretrieve(
                        'http://biogeo.ucdavis.edu/data/gadm2.8/shp/' +
                        zip_file_name,
                        dirpath + '/zipfiles/' + zip_file_name)

            except ContentTooShortError as e:
                print (e.__str__())
            except Exception as e:
                print (e.__str__())
        return dirpath

    def unzip_gadm_files(self, dirpath):
        '''
            Function is used to unzip the shape files zip folder and
            save it to one step up in the hierarchy.
        '''

        for fl in self.file_names:
            with open(dirpath + '/zipfiles/' + fl + '.zip'):
                zippedFiles = ZipFile(dirpath + '/zipfiles/' + fl + '.zip')
                zippedFiles.extractall(dirpath + '/' + fl)

    def create_topojson(self, dirpath):
        '''
            Function calls 'topojson_generator()' to generate topojson
            files using shape files'
        '''

        for dr in self.file_names:
            # creating generator for generating topojson file
            shp_dir = dirpath + '/' + dr
            for i in self.topojson_generator(shp_dir):
                print (i)

    def topojson_generator(self, shp_dir):
        '''
            Function is used to generate topojson files using shape files.
            It passes the topojson files to external file named 'shape.py'
            to generate excel-maps using these files.
        '''

        chdir(shp_dir)
        file_str = ''

        for f in listdir(shp_dir):
            # only targetting files of .shp extension'
            if f.endswith('.shp'):
                file_str += f + ' '
        file_str = file_str[:len(file_str) - 1]

        if file_str:
            cmd = 'topojson' + ' ' + file_str + ' ' + '  -o ' + \
                path.basename(shp_dir) + '.json'

            process = Popen(shlex.split(cmd), shell=True)

            try:
                process.wait()
                chdir(self.dirname)

                cmd = 'python shape.py ' + shp_dir + \
                    '/' + path.basename(shp_dir) + '.json'

                process = Popen(cmd, shell=True)
                process.wait()
            except:
                print ('Timout for executing the process.')
                process.kill()
            finally:
                chdir(dirpath)
        yield file_str


if __name__ == '__main__':
    gsf = GetShapeFiles()
    dirpath = gsf.gadm_download_files()
    gsf.unzip_gadm_files(dirpath)
    gsf.create_topojson(dirpath)
