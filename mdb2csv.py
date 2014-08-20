#!/usr/bin/python

import os
import csv
import ctypes
import datetime
import decimal
import pypyodbc
import random

ODBC_ADD_DSN = 1
ODBC_CONFIG_DSN = 2
ODBC_REMOVE_DSN = 3
ODBC_ADD_SYS_DSN = 4
ODBC_CONFIG_SYS_DSN = 5
ODBC_REMOVE_SYS_DSN = 6

def get_mdb_driver():
    mdb_driver = [d for d in pypyodbc.drivers() if 'Microsoft Access Driver (*.mdb' in d]
    return mdb_driver[0]

def edit_dsn(driver, change='add', **kw):
    nul = chr(0)
    attributes = []
    for attr in kw.keys():
        attributes.append('%s=%s' % (attr, kw[attr]))
    attrs = nul.join(attributes) + nul + nul
    driver = driver.encode('mbcs')
    attrs = attrs.encode('mbcs')
    if change == 'add':
        change = ODBC_ADD_DSN
    else:
        change = ODBC_REMOVE_DSN
    return ctypes.windll.ODBCCP32.SQLConfigDataSource(0, change, driver, attrs)

def dsn_mdb2csv(connection_string, destdir):
    connection = pypyodbc.connect(connection_string)
    cur = connection.cursor()
    
    f = lambda x: os.path.join(destdir, x)
    if not os.path.exists(destdir):
        os.mkdir(destdir)
    
    cur.tables()
    table_list = cur.fetchall()
    
    for tbl in table_list:
        tname = tbl[2]
        ttype = tbl[3]
        if ttype == 'TABLE':
            
            fd = open(f(tname.lower() + '.csv'), 'wb')
            fds = open(f(tname.lower() + '.create.sql'), 'wb')
            
            w = csv.writer(fd)
            print 'Dumping', tname
            cur.execute('SELECT * FROM %s' % tname)
            # add column names
            w.writerow([i[0] for i in cur.description])
            
            
            fields = []
            for i in cur.description:
                
                ftype = {
                    datetime.datetime: 'timestamptz',
                    unicode: 'varchar',
                    decimal.Decimal: 'numeric',
                    int: 'integer'
                }[i[1]]
                 
                if ftype == 'varchar':
                    ftype = '%s(%s)' % (ftype, max(i[2], i[3], i[4]))
                elif ftype == 'numeric':
                    ftype = '%s(%s, %s)' % (ftype, max(i[2], i[3], i[4]), i[5])
                if i[6] == False:
                    ftype = ftype + ' NOT NULL'
                fields.append('"%s" %s' %  (i[0], ftype))
            
            
            fds.write(('CREATE TABLE "%s" (\n\t' % (tname,)) + ',\n\t'.join(fields) + '\n);');
            fds.close()
            
            for rec in cur:
                row = []
                for i in rec:
                    if isinstance(i, basestring):
                        row.append(i.encode('utf-8'))
                    else:
                        row.append(i)
                w.writerow(row)
            
            fd.close()
    connection.close()

def mdb2csv(filename, destdir):
    n = random.randint(1, 2000)
    
    dsn = "mdb2csv_%d" % (n,)
    params = {
        'DSN': dsn,
        'DBQ': filename,
        'READONLY': '1',
    }
    
    
    driver = get_mdb_driver()
    print 'adding', dsn
    created = edit_dsn(driver, 'add', **params)
    if not created:
        print 'error creating dsn'
        return

    # do stuff
    connection_string = 'DSN=%s' % dsn
    dsn_mdb2csv(connection_string, destdir)

    removed = edit_dsn(driver, 'del', **params)
    if not removed:
        print 'error removing dsn'


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description='Convert Access MDB to CSV.')
    parser.add_argument('filename', help='MDB database file')
    parser.add_argument('destdir', help='destination directory for CSV files (one per table)')

    args = parser.parse_args()
    print args
    
    fn = args.filename
    destdir = args.destdir

    mdb2csv(fn, destdir)
