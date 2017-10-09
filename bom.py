#!/usr/bin/env python
# This script converts KiCad files into an Microsoft Excel table for
# a Bill of Materials.
# Parts with identical Manufacturer Part Numbers (mpn) are combined,
# independent of the part reference
# Written by Martin Hepp, www.heppnetz.de

import os
import sys
import argparse
import logging
import xml.etree.ElementTree as ET
import pandas as pd


def main(args, loglevel):
    logging.basicConfig(format='%(levelname)s: %(message)s', level=loglevel)
    if not os.path.exists(args.filename):
        print('ERROR: No such input file: %s' % args.filename)
        sys.exit()
    target = os.path.splitext(args.filename)[0] + '.xlsx'
    if os.path.exists(target):
        yes = {'yes', 'y', 'ye', ''}
        choice = input('''Target filename %s exists. \
            Overwrite? [y/n]''' % target).lower()
        if choice not in yes:
            print('Operation cancelled.')
            sys.exit()
    logging.info('Processing %s' % args.filename)
    if args.mpn is not None:
        MPN_FIELD = args.mpn
    else:
        MPN_FIELD = 'MPN'

    logging.info('Using %s as the MPN field name' % MPN_FIELD)
    bom = {}
    tree = ET.parse(args.filename)
    root = tree.getroot()
    c = root.find('components')
    for comp in c.findall('comp'):
        component = {}
        component['ref'] = comp.get('ref', 'missing').strip()
        component['qtty'] = 1
        if comp.find('value') is not None:
            component['value'] = comp.find('value').text
        else:
            component['value'] = 'missing'
            logging.warning('Value missing for  %s' % component['ref'])
        if comp.find('footprint') is not None:
            component['footprint'] = comp.find('footprint').text
        else:
            component['footprint'] = 'missing'
            logging.warning('Footprint missing for  %s' % component['ref'])
        fields = comp.find('fields')
        if fields is not None:
            for field in fields.findall('field'):
                name = field.get('name')
                value = field.text
                component[name] = value
            mpn = component.get(MPN_FIELD, '').strip()
            if not mpn or mpn == 'n/a':
                mpn = 'N/A-' + component['ref']
        else:
            mpn = 'N/A-' + component['ref']
        if bom.get(mpn, None) is not None:
            bom[mpn]['qtty'] += 1
            bom[mpn]['ref'] += ', ' + component['ref']
        else:
            bom[mpn] = component

    logging.info('Writing results to %s' % target)
    df = pd.DataFrame.from_dict(bom, orient='index')
    df = df.sort_values('ref')
    writer = pd.ExcelWriter(target)
    df.to_excel(writer, 'BOM', index=False)
    writer.save()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='''Creates an MS Excel\
     table from the given KiCad BOM file in XML.''')
    parser.add_argument('filename',
                        help='The filename of the KiCad XML file')
    parser.add_argument("--mpn", help='''Set the KiCad field name for\
    the manufacturer part number (MPN).''')
    args = parser.parse_args()
    loglevel = logging.INFO
    main(args, loglevel)
