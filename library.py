#! /usr/bin/env python3
import csv
import os
import re
from argparse import ArgumentParser
from dataclasses import dataclass

from olefile.olefile import OleStream

import altium
from sys import argv, stdout
from warnings import warn
import configparser


class LibPkg(object):

    def __init__(self, file) -> None:
        super().__init__()
        self.dir = os.path.dirname(file)
        self.name = os.path.basename(file)
        config = configparser.ConfigParser()
        config.read(file, encoding='utf-8-sig')
        self.design = config['Design']
        if self.design['version'] != '1.0':
            raise OSError('Invalid LibPkg version')
        self.documents = [config[x] for x in config.sections() if x.startswith('Document')]
        self.schlibs = {}
        for doc in self.documents:
            documentpath = doc['documentpath']
            path = os.path.join(self.dir, documentpath)
            if os.path.splitext(path)[-1].lower() == '.schlib':
                print(path)
                self.schlibs[os.path.basename(documentpath)] = SchLib(path)

    def parts_to_csv(self, file, fieldnames):
        with open(file, 'w', newline='') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames, extrasaction='ignore')

            writer.writeheader()
            for schlib in self.schlibs.values():
                writer.writerows([part.params() for key, part in schlib.get_parts().items()])


    def list_documents(self):
        return self.documents


class SchPart(object):

    def __init__(self, stream: OleStream, description='') -> None:
        super().__init__()
        records = altium.iter_records(stream)
        records = (altium.parse_properties(stream, record) for record in records)

        # Decode header
        header = next(records)
        self.desc = description
        self.id = header.get('DESIGNITEMID')
        if self.id is not None:
            self.id = self.id.decode('ascii')
        else:
            self.id = header.get('LIBREFERENCE').decode('ascii')

        # Decode properties and footprints
        self.designator = ''
        self.properties = {}
        self.footprints = {}
        for record in records:
            if record is None:
                continue

            if record.get_int('RECORD') == 41:
                # Property
                name = record.get('NAME').decode('ascii')
                text = record.get('TEXT', b'').decode('ascii')
                self.properties[name] = text
            elif record.get_int('RECORD') == 34:
                # Designator
                self.designator = record.get('TEXT').decode('ascii')
            elif record.get_int('RECORD') == 45:
                # Footprint
                name = record.get('MODELNAME').decode('ascii')
                text = record.get('DESCRIPTION', b'').decode('ascii')
                self.footprints[name] = text
            #else:
            #    print(record)

    def params(self):
        params = {
            'id': self.id,
            'designator': self.designator,
            'description': self.desc,
        }
        params.update(self.properties)
        return params

    def parts_to_csv(self, file, fieldnames):
        with open(file, 'w', newline='') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames, extrasaction='ignore')

            writer.writeheader()
            writer.writerows(self.params())

    def __str__(self) -> str:
        return '{}, {}, {}, {}'.format(self.id, self.designator, self.desc, self.properties)


class SchLib(object):

    def __init__(self, file) -> None:
        super().__init__()
        file = altium.OleFileIO(file)
        stream = file.openstream('FileHeader')
        records = altium.iter_records(stream)
        records = (altium.parse_properties(stream, record) for record in records)
        # Parse header
        header = next(records)
        header.check("HEADER", b"Protel for Windows - Schematic Library Editor Binary File Version 5.0")
        header.check("MINORVERSION", None, b"2")

        # Schematic parts
        self.parts = {}
        cnt = 0
        while True:
            libref = header.get("LIBREF{}".format(cnt), None)
            desc = header.get("COMPDESCR{}".format(cnt), b'')\
                .decode('utf8', errors='replace')\
                .replace('\n', ' ')
            #print(desc)
            if not libref:
                break
            stream = file.openstream([libref.decode('ascii'), 'Data'])
            try:
                part = SchPart(stream, description=desc)
                self.parts[libref.decode('ascii')] = part
            except ValueError as ve:
                pass
            cnt += 1

    def get_parts(self):
        return self.parts


if __name__ == "__main__":
    parser = ArgumentParser(description="Extract parts from .SchLib or .LibPkg")
    parser.add_argument("file")
    parser.add_argument("fields", nargs='*')
    parser.add_argument("--output", required=False)
    args = parser.parse_args()

    fieldnames = ['id', 'designator', 'description', 'Comment']
    fieldnames.extend(args.fields)

    if args.output is None:
        args.output = os.path.basename(args.file) + '.csv'

    ext = os.path.splitext(args.file)[-1].lower()
    if ext == '.libpkg':
        libpkg = LibPkg(args.file)
        libpkg.parts_to_csv(args.output, fieldnames)
    elif ext == '.schlib':
        sch = SchLib(args.file)
    else:
        raise OSError('Unknown file format \'{}\''.format(ext))