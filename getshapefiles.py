import os
import glob
import logging
import requests
import lxml.html
import subprocess
from six.moves.urllib.request import urlretrieve
from zipfile import ZipFile
# from six.moves.urllib.error import ContentTooShortError
import shape
import shlex
import argparse


class JSONFileObject(object):
    encoding = 'utf-8'
    file = []
    key = []


def gadm_download_files(target, limit=10):
    '''
    Download the shape files from gadm into target folder with the following
    structure:

    - `/target/zipfiles` stores the downloaded ZIP files
    - `/target/AFG_adm_shp/` stores the topojson files for AFG_ADM.shp
    - `/target/...` etc
    '''
    gadm_page_url = 'http://www.gadm.org/country'
    zip_dir = os.path.join(target, 'gadmzips/zipfiles')
    if not os.path.exists(zip_dir):
        os.makedirs(zip_dir)

    response = requests.get(gadm_page_url)
    tree = lxml.html.fromstring(response.content)
    country_codes = tree.xpath('//select[@name="cnt"]/option')
    for country_code in country_codes[:limit]:
        option_value = country_code.get('value')
        if option_value is None:
            logging.warn('Skipping %s', lxml.html.tostring(country_code))
            continue
        raw_file_name = option_value.split('_')[0] + '_adm_shp'
        zip_name = raw_file_name + '.zip'
        zip_path = os.path.join(zip_dir, zip_name)
        if not os.path.exists(zip_path):
            logging.info('%s: downloading', zip_name)
            urlretrieve(
                'http://biogeo.ucdavis.edu/data/gadm2.8/shp/' +
                zip_name, zip_path)
        else:
            logging.info('%s: downloaded', zip_name)
        yield zip_path


def unzip_gadm_file(zip_path):
    '''
    Unzip zip_path='E:\cartogram\gadmzips/zipfiles/ATA_adm_shp.zip'
    into 'gadmzips/ATA_adm_shp'
    '''
    dirname, filename = os.path.split(
        zip_path)         # E:\cartogram\gadmzips/zipfiles, ATA_adm_shp.zip
    folder = os.path.splitext(filename)[0]              # ATA_adm_shp
    parent = os.path.dirname(dirname)                   # gadmzips
    shapefile_dir = os.path.join(parent, folder)        # gadmzips/ATA_adm_shp
    if not os.path.isdir(shapefile_dir):
        logging.info('%s: extracting', shapefile_dir)
        ZipFile(zip_path).extractall(shapefile_dir)
    return shapefile_dir


def create_topojson(shp_dir, json_obj):
    '''
    Generate topojson files using shape files.
    It passes the topojson files to external file named 'shape.py'
    to generate excel-maps using these files.
    '''
    for shapefile_path in glob.glob(os.path.join(shp_dir, '*.shp')):
        subdir, shapefile_name = os.path.split(os.path.abspath(shapefile_path))
        json_file = os.path.basename(shapefile_name) + '.json'
        if not os.path.exists(os.path.join(subdir, json_file)):
            logging.info('%s: creating', json_file)
            cmd = 'topojson ' + shapefile_name + ' -o ' + json_file
            process = subprocess.Popen(
                shlex.split(cmd),
                cwd=subdir, shell=True)
            process.wait()
            json_obj.file = [json_file]
            shape.main(json_obj)
        else:
            logging.info('%s: exists', json_file)


if __name__ == '__main__':
    # setting up logging level
    logging.basicConfig(level=logging.INFO)

    # argument parser to specify the directory location
    parser = argparse.ArgumentParser()
    parser.add_argument(
        '-d',
        '--directory',
        help='directory path inside which zipfiles should be downloaded',
        default=os.getcwd())
    args = parser.parse_args()
    logging.info('%s: creating directory structure', args.directory)

    # Object for shape module working
    json_obj = JSONFileObject()
    for zip_path in gadm_download_files(target=os.path.abspath(args.directory),
                                        limit=1):
        shapefile_dir = unzip_gadm_file(zip_path)
        create_topojson(shapefile_dir, json_obj)
