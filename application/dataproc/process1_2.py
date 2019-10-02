# coding=utf-8

from pymongo import MongoClient


def process1_2(process):
    client = MongoClient(host='127.0.0.1', port=27017)
    db = client['emr']
    col_from = db[process['from_collection']]
    col_to = db[process['to_collection']]

    if process['drop_collections_before_save'].upper() == 'YES':
        col_to.delete_many({})

    documents = col_from.find()

    for doc in documents:
        data = {}
        data['YY'] = '浙四医院'
        data['XM'] = doc['入院记录']['姓名']
        data['XB'] = doc['入院记录']['性别']
        data['CSNY'] = doc['入院记录']['出生日期']
        data['RYSJ'] = doc['入院记录']['入院时间']
        data['BLH'] = doc['首次病程记录']['病历号']
        col_to.insert_one(data)

    client.close()
