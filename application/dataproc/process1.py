# coding=utf-8

import os
import sys
from typing import Dict, Any

import docx
import re


def process1_1(process):
    for dir in process['source_dir']:
        print("Reading all files from {}".format(dir))
        files = os.listdir(dir)
        #files = ['56444.docx']
        for file in files:
            if not file.endswith('.docx'):
                continue
            print("Reading {}".format(file))
            read_docx(os.path.join(dir, file))


def read_docx(file_path):
    doc = docx.Document(file_path)
    full_text = ''
    for para in doc.paragraphs:
        full_text = full_text + para.text + '\n'

    # 将full_text分成三部份admission_record_text、progress_note_text、discharge_record_text

    regex ='(入  院  记  录.*?)(首次病程记录|出  院  记  录)'
    r = re.search(regex, full_text, re.S)
    if r is not None:
        admission_record_text = r.group(1)
    else:
        regex = '入  院  记  录.*'
        r = re.search(regex, full_text, re.S)
        admission_record_text = r.group() if r is not None else ''

    regex ='(首次病程记录.*?)(入  院  记  录|出  院  记  录)'
    r = re.search(regex, full_text, re.S)
    if r is not None:
        progress_note_text = r.group(1)
    else:
        regex = '首次病程记录.*'
        r = re.search(regex, full_text, re.S)
        progress_note_text = r.group() if r is not None else ''

    regex ='(出  院  记  录.*?)(入  院  记  录|首次病程记录)'
    r = re.search(regex, full_text, re.S)
    if r is not None:
        discharge_record_text = r.group(1)
    else:
        regex ='出  院  记  录.*'
        r = re.search(regex, full_text, re.S)
        discharge_record_text = r.group() if r is not None else ''

    regex = '入  院  记  录.*出  院  记  录'
    r = re.search(regex, full_text, re.S)
    if r is not None:
        t1 = 0
        t2 = 1
    else:
        t1 = 1
        t2 = 0

    admission_record = dict()
    progress_note = dict()
    discharge_record = dict()

    table = doc.tables[t1]
    admission_record['姓名'] = table.cell(0, 1).text
    admission_record['职业'] = table.cell(0, 3).text
    admission_record['性别'] = table.cell(1, 1).text
    admission_record['工作单位'] = table.cell(1, 3).text
    admission_record['出生日期'] = table.cell(2, 1).text
    admission_record['户口地址'] = table.cell(2, 3).text
    admission_record['婚姻'] = table.cell(3, 1).text
    admission_record['联系电话'] = table.cell(3, 3).text
    admission_record['出生地'] = table.cell(4, 1).text
    admission_record['入院时间'] = table.cell(4, 3).text
    admission_record['民族'] = table.cell(5, 1).text
    admission_record['病史陈述者'] = table.cell(5, 3).text

    table = doc.tables[t2]
    i = 0 if len(table.rows) == 3 else 1
    discharge_record['入院日期'] = table.cell(i, 1).text
    discharge_record['出院日期'] = table.cell(i, 3).text
    discharge_record['入院诊断'] = table.cell(i + 1, 1).text
    discharge_record['出院诊断'] = table.cell(i + 1, 3).text
    discharge_record['住院天数'] = table.cell(i + 2, 1).text

    # 处理入院记录

    regex = '主诉\(Chief complaint\)：(?P<主诉>.+)\n' \
            '现病史（History of Present Illness）:(?P<现病史>.+)\n' \
            '既往史（Past History）:(?P<既往史>.+)\n' \
            '目前使用的药物（At Present The Drugs）：（含我院用药情况及患者提供的用药情况）(?P<目前使用药物>.+)\n' \
            '成瘾药物\(Drug Addiction\):(?P<成瘾药物>.+)\n' \
            '个人史（Personal History）:(?P<个人史>.+?)\n' \
            '(?P<menstrual>.*)' \
            '婚育史（Obstetrical History）:(?P<婚育史>.+)\n' \
            '家族史（Family History）:(?P<家族史>.+)\n' \
            '体格检查（Physical Examination）：(?P<体格检查>.+)\n' \
            '辅助检查（Diagnostic Examination）：(?P<辅助检查>.+)\n' \
            '营养风险筛查\(Nutritional Assessment\).+体重指数\(BMI\):(?P<体重指数>.+)\n' \
            '疾病相关评分:\n(?P<疾病相关评分>.+)\n' \
            '营养受损评分:\n(?P<营养受损评分>.+)\n' \
            '年龄评分:(?P<年龄评分>.+)\n' \
            '营养风险评分:(?P<营养风险评分>.+?分).*\n' \
            '是否请营养科会诊:(?P<是否请营养科会诊>.+)\n' \
            '功能评估:\(Function  Accessment\)\n((入院ADL评分:(?P<入院ADL评分>.*))|(入院ADL评分分级:(?P<入院ADL评分分级>.*)))\n' \
            '是否请康复科会诊:(?P<是否请康复科会诊>.+)\n' \
            '心理评估\(Psychological Assessment\)\n护理入院心理评估是否阳性:(?P<护理入院心理评估是否阳性>.*)\n' \
            '是否请心理卫生科会诊:(?P<是否请心理卫生科会诊>.*)\n初步诊断'

    match = re.search(regex, admission_record_text, re.S)
    a = match.groupdict()
    admission_record['主诉'] = a['主诉']
    admission_record['现病史'] = a['现病史']
    admission_record['既往史'] = a['既往史']
    admission_record['目前使用药物'] = a['目前使用药物']
    admission_record['成瘾药物'] = a['成瘾药物']
    admission_record['个人史'] = a['个人史']
    admission_record['婚育史'] = a['婚育史']
    admission_record['家族史'] = a['家族史']
    admission_record['体格检查'] = a['体格检查']
    admission_record['辅助检查'] = a['辅助检查']
    admission_record['营养风险筛查'] = {}
    admission_record['营养风险筛查']['体重指数'] = a['体重指数']
    admission_record['营养风险筛查']['疾病相关评分'] = a['疾病相关评分']
    admission_record['营养风险筛查']['营养受损评分'] = a['营养受损评分']
    admission_record['营养风险筛查']['年龄评分'] = a['年龄评分']
    admission_record['营养风险筛查']['营养风险评分'] = a['营养风险评分']
    admission_record['营养风险筛查']['是否请营养科会诊'] = a['是否请营养科会诊']
    admission_record['功能评估'] = {}
    if '入院ADL评分' in a.keys():
        admission_record['功能评估']['入院ADL评分'] = a['入院ADL评分']
    elif '入院ADL评分分级' in a.keys():
        admission_record['功能评估']['入院ADL评分'] = a['入院ADL评分分级']
    else:
        admission_record['功能评估']['入院ADL评分'] = ''

    admission_record['功能评估']['是否请康复科会诊'] = a['是否请康复科会诊']
    admission_record['心理评估'] = {}
    admission_record['心理评估']['护理入院心理评估是否阳性'] = a['护理入院心理评估是否阳性']
    admission_record['心理评估']['是否请心理卫生科会诊'] = a['是否请心理卫生科会诊']

    regex = '月经史（Menstrual History）:(.+)\n'
    r = re.search(regex, a['menstrual'], re.S)
    admission_record['月经史'] = r.group(1) if r is not None else ''

    regex = '初步诊断\(Diagnosis\).*'
    r = re.search(regex, admission_record_text, re.S)
    diagnosis = r.group()

    regex = '初步诊断\(Diagnosis\)：\n(.+?)医师签名：'
    r = re.search(regex, diagnosis, re.S)
    admission_record['初步诊断'] = r.group(1) if r is not None else ''

    regex = '修正诊断\(Diagnosis\)：\n(.+?)医生签名：'
    r = re.search(regex, diagnosis, re.S)
    admission_record['修正诊断'] = r.group(1) if r is not None else ''

    regex = '补充诊断\(Diagnosis\)：\n(.+?)医生签名：'
    r = re.search(regex, diagnosis, re.S)
    admission_record['补充诊断'] = r.group(1) if r is not None else ''

    # 处理首次病程记录

    regex = '病例特点： \n(?P<病例特点>.+)\n' \
            '初步诊断：(?P<初步诊断1>.+)\n' \
            '诊断依据：(?P<诊断依据>.+)\n' \
            '鉴别诊断：(?P<鉴别诊断>.+)\n' \
            '诊疗计划：\n(?P<诊疗计划>.+)\n.+(医师签名：|记录医生：)'

    match = re.search(regex, progress_note_text, re.S)
    a = match.groupdict()

    progress_note['病例特点'] = a['病例特点']
    progress_note['初步诊断'] = a['初步诊断1']
    progress_note['诊断依据'] = a['诊断依据']
    progress_note['鉴别诊断'] = a['鉴别诊断']
    progress_note['诊疗计划'] = a['诊疗计划']

    # 处理出院记录

    regex = '入院情况:(?P<入院情况>.+)\n' \
            '住院经过:(?P<住院经过>.+)\n' \
            '出院情况:(?P<出院情况>.+)\n' \
            '出院医嘱(:|：)(?P<出院医嘱>.+)\n' \
            '健康教育:(?P<健康教育>.+)\n' \
            '随访计划:(?P<随访计划>.+)医师签名：'

    match = re.search(regex, discharge_record_text, re.S)
    a = match.groupdict()

    discharge_record['入院情况'] = a['入院情况']
    discharge_record['住院经过'] = a['住院经过']
    discharge_record['出院情况'] = a['出院情况']
    discharge_record['出院医嘱'] = a['出院医嘱']
    discharge_record['健康教育'] = a['健康教育']
    discharge_record['随访计划'] = a['随访计划']

    # 合并入院记录、首次病程记录、出院记录到medical_record中

    medical_record = dict()
    medical_record['入院记录'] = admission_record
    medical_record['首次病程记录'] = progress_note
    medical_record['出院记录'] = discharge_record

    trim_dict_values(medical_record)
    trim_dict_values(medical_record['入院记录'])
    trim_dict_values(medical_record['首次病程记录'])
    trim_dict_values(medical_record['出院记录'])

    pass


def trim_dict_values(dic):
    for key in dic:
        value = dic[key]
        if isinstance(value, str):
            dic[key] = value.strip(' \n')
