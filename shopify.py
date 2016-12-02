'''
Extracts properties from a TopoJSON file into a Shopify CSV file
'''

import io
import json
import pandas as pd
from tqdm import tqdm
from six import StringIO


def properties(path, encoding='utf-8'):
    with io.open(path, encoding=encoding) as handle:
        topo = json.load(handle)
    data = []
    for shape in topo['objects'].values():
        for geom in shape['geometries']:
            data.append(geom.get('properties', {}))
    return pd.DataFrame(data)


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(
        description=__doc__.strip(),
        formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('-o', '--output', help='Output CSV file', default='product-template.csv')
    parser.add_argument('file', help='TopoJSON file', nargs='+')
    args = parser.parse_args()

    result = []
    for path in tqdm(args.file):
        data = properties(path)

        buf = StringIO()
        data.to_html(buf, index=False, classes=None)

        title = path
        unique_id = path
        info = {
            'Handle': unique_id,
            'Title': title,
            'Body (HTML)': buf.getvalue(),
            'Variant Price': 25000
        }
        result.append(info)

    result = pd.DataFrame(result)
    result.to_csv(args.output, index=False, encoding='utf-8')
