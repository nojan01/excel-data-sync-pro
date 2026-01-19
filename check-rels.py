#!/usr/bin/env python3
"""Prüft XLSX auf absolute Pfade in Relationships"""
import zipfile
import os
import tempfile
import shutil
import re

export_path = '/Users/nojan/Desktop/Export.xlsx'

# Temporäres Verzeichnis
temp_dir = tempfile.mkdtemp()
print(f'Temp dir: {temp_dir}')

# Extrahiere die XLSX
with zipfile.ZipFile(export_path, 'r') as zf:
    zf.extractall(temp_dir)

# Lies alle .rels Dateien und zeige absolute Pfade
for root, dirs, files in os.walk(temp_dir):
    for f in files:
        if f.endswith('.rels'):
            full_path = os.path.join(root, f)
            rel_path = full_path.replace(temp_dir + '/', '')
            with open(full_path, 'r') as fp:
                content = fp.read()
            # Finde alle absoluten Pfade (Target beginnt mit /)
            abs_paths = re.findall(r'Target="(/[^"]+)"', content)
            if abs_paths:
                print(f'{rel_path}:')
                for p in abs_paths:
                    print(f'  ABSOLUTE: {p}')

# Cleanup
shutil.rmtree(temp_dir)
print('Done.')
