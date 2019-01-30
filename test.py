# coding=gbk
# coding: utf-8
import xlsxwriter

dict=[{'name': '谌宁生', 'introduction': '谌宁生 主任医师，教授，硕士生导师，第二批国家级老中医药专家学术经验继承工作指导老师 全国名老中医药专家 医院首届名医;国家肝病中医医疗中心学术奠基人，从医60余年，对肝病有深入研究，擅治各种急慢性肝炎、肝硬化、重型肝炎、脂肪肝及内科疑难杂症，对肝癌亦有深入的研究。历任多个科室主任，主持承担国家“八五”科技攻关、湖南省科委、教委及卫生厅局等肝病科研重点课题多项，获湖南省科技进步奖2次，湖南省中医药科技进步奖3次。担任中华中医药学会终身理事、国家自然科学基金评审委员、世界教科文卫组织专家成员、国家中医药管理局“十一五”中医肝病重点专科协作组专家学术指导委员会委员、国家重大专项“十二五”重肝课题方案专家论证会特邀专家。担任十余家期刊杂志副总编、副主编、常务编委、编委及特约撰稿人，在国内外50多家期刊、杂志、书报发表论文150余篇，其中10余篇获国际优秀论文奖及金杯奖;主编著作《中医治疗病毒性肝炎的研究与实践》、副主编4部、特约编委7部、参编13部。获得“2000千年名医”、“环球时代杰出人物”、“海内外杰出爱国人士”、“共和国杰出人物”、“中国优秀医学专家”等多个荣誉称号;多次应邀至台湾、韩国等地进行讲学及医学访问，2013年4月受邀赴泰国参加国际传统医学与养生大会，受泰国亲皇亲切接见。', 'title': '主任医师 教授 硕士生导师 第二批国家级老中医药专家学术经验继承工作指导老师 全国名老中医药专家 医院首届名医', 'imgUrl': '1111'}]

def generate_excel(rec_data):
    workbook = xlsxwriter.Workbook('C://Users//Hey//Desktop//emp.xlsx')
    worksheet = workbook.add_worksheet()

    # 设定格式，等号左边格式名称自定义，字典中格式为指定选项
     # bold：加粗，num_format:数字格式
    bold_format = workbook.add_format({'bold': True})
    money_format = workbook.add_format({'num_format': '$#,##0'})
    date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})
    # 用符号标记位置，例如：A列1行
    worksheet.write('A1', 'name', bold_format)
    worksheet.write('B1', 'introduction', bold_format)
    worksheet.write('C1', 'title', bold_format)
    worksheet.write('D1', 'imgUrl', bold_format)
    row = 1
    col = 0
    for item in rec_data:
        worksheet.write_string(row, col, item['name'])
        worksheet.write_string(row, col + 1, item['introduction'])
        worksheet.write_string(row, col + 2, str(item['title']))
        worksheet.write_string(row, col + 3, item['imgUrl'])
        row += 1
        workbook.close()
generate_excel(dict)