import os
import glob
import logging
import requests
import lxml.html
import subprocess
from six.moves.urllib.request import urlretrieve
from zipfile import ZipFile
from six.moves.urllib.error import ContentTooShortError


def gadm_download_files(target='gadmzips', limit=10):
    '''
    Download the shape files from gadm into target folder with the following
    structure:

    - `/target/zipfiles` stores the downloaded ZIP files
    - `/target/AFG_adm_shp/` stores the topojson files for AFG_ADM.shp
    - `/target/...` etc
    '''
    zip_dir = os.path.join(target, 'zipfiles')
    if not os.path.exists(zip_dir):
        makedirs(zip_dir)

    gadm_page_url = 'http://www.gadm.org/country'
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
            urlretrieve('http://biogeo.ucdavis.edu/data/gadm2.8/shp/' + zip_name, zip_path)
        else:
            logging.info('%s: downloaded', zip_name)
        yield zip_path


def unzip_gadm_file(zip_path):
    '''
    Unzip zip_path='gadmzips/zipfiles/ATA_adm_shp.zip' into 'gadmzips/ATA_adm_shp'
    '''
    dirname, filename = os.path.split(zip_path)         # gadmzips/zipfiles, ATA_adm_shp.zip
    folder = os.path.splitext(filename)[0]              # ATA_adm_shp
    parent = os.path.dirname(dirname)                   # gadmzips
    shapefile_dir = os.path.join(parent, folder)        # gadmzips/ATA_adm_shp
    if not os.path.isdir(shapefile_dir):
        logging.info('%s: extracting', shapefile_dir)
        ZipFile(zip_path).extractall(shapefile_dir)
    return shapefile_dir


def create_topojson(shp_dir):
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
            returncode = subprocess.call(
                cmd['topojson', shapefile_name, '-o', json_file],
                cwd=subdir, shell=True)
        else:
            logging.info('%s: exists', json_file)


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    for zip_path in gadm_download_files(limit=10):
        shapefile_dir = unzip_gadm_file(zip_path)
        create_topojson(shapefile_dir)
