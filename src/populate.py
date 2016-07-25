#!/usr/bin/env python
#coding:utf-8
"""
  Author:   --<>
  Purpose: 
  Created: 2016/3/29
"""

from dataset import *
import random, string

def createRandomStrings(l,n):
    """create list of l random strings, each of length n"""
    names = []
    for i in range(l):
        val = ''.join(random.choice(string.ascii_lowercase) for x in range(n))
        names.append(val)
    return names

def createData(rows=20, cols=5):
    """Creare random dict for test data"""

    data = {}
    names = createRandomStrings(rows,16)
    colnames = createRandomStrings(cols,5)
    for n in names:
        data[n]={}
        data[n]['label'] = n
    for c in range(0,cols):
        colname=colnames[c]
        vals = [round(random.normalvariate(100,50),2) for i in range(0,len(names))]
        vals = sorted(vals)
        i=0
        for n in names:
            data[n][colname] = vals[i]
            i+=1
    return data

def init_database():
    mbom_db.connect()
    mbom_db.create_tables([id_generator, mat_basic_info, mat_extra_info, mat_info, bom_header,bom_item, prj_bom_link])
    #nstd_mat_fin.get(nstd_mat_fin.mat_no=='330172045')
    #nstd_mat_fin.delete_instance(nstd_mat_table)
    
def insert_data_into_tables():
    q = id_generator.insert(desc='BOM header ID取号',step=1,current=1, pre_character='BH')
    q.execute()

def close_database():
    mbom_db.close()
    
if __name__ == '__main__':
    init_database()
    insert_data_into_tables()
    close_database()