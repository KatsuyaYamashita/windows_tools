#
# CSVファイルを読んで書き出すスクリプト
# %proc_csv orig_fname
#
# 1. rename orig_fname -> backup_fname
# 2. csvファイルを1レコードずつ読み出して別ファイルに書き出す
#     backup_fname -> orig_fname
#

import sys
import os
import csv

def replace_fileext(path, ext):
    orig_root, orig_ext = os.path.splitext(path)
    backup_fname = orig_root + ext
    return backup_fname

def replace_filename(path, fname):
    root = os.path.dirname(path)
    return root + '\\' + fname

def backup_file(orig_fname, backup_fname):
    os.rename(orig_fname, backup_fname)

def process_csvfile(input_fname, output_fname):
    try:
        rfile = open(input_fname)
        wfile = open(output_fname, 'w', newline='')
    except OSError as err:
        print("read error: {0}".format(err))
    else:
        sr = csv.reader(rfile)
        sw = csv.writer(wfile, quoting=csv.QUOTE_ALL)
        total = 0
        for row in sr:
            wrow = []
            for column in row:
                wrow.append(column)
                
            sw.writerow(wrow)
            total += 1

        rfile.close()
        wfile.close()
                
    print("total=%d" % total, file=sys.stderr)


if __name__ == '__main__':
    argvs = sys.argv
    argc = len(argvs)
    if (argc < 2):
        print("Invalid Argument", file=sys.stderr)
    else:
        input_fname = argvs[1]
        output_fname = argvs[1]
        backup_fname = replace_fileext(input_fname, '.bak')
        error_fname = replace_filename(input_fname, 'error.txt')

        print("input=%s,backup=%s" % (input_fname, backup_fname))

        backup_file(input_fname, backup_fname)
        process_csvfile(backup_fname, output_fname)
