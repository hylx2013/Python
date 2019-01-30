# coding=gbk
# coding: utf-8
import xlsxwriter

dict=[{'name': '������', 'introduction': '������ ����ҽʦ�����ڣ�˶ʿ����ʦ���ڶ������Ҽ�����ҽҩר��ѧ������̳й���ָ����ʦ ȫ��������ҽҩר�� ҽԺ�׽���ҽ;���Ҹβ���ҽҽ������ѧ������ˣ���ҽ60���꣬�Ըβ��������о������θ��ּ����Ը��ס���Ӳ�������͸��ס�֬���μ��ڿ�������֢���Ըΰ�����������о������ζ���������Σ����ֳе����ҡ����塱�Ƽ����ء�����ʡ��ί����ί���������ֵȸβ������ص�����������ʡ�Ƽ�������2�Σ�����ʡ��ҽҩ�Ƽ�������3�Ρ������л���ҽҩѧ���������¡�������Ȼ��ѧ��������ίԱ������̿�������֯ר�ҳ�Ա��������ҽҩ����֡�ʮһ�塱��ҽ�β��ص�ר��Э����ר��ѧ��ָ��ίԱ��ίԱ�������ش�ר�ʮ���塱�ظο��ⷽ��ר����֤������ר�ҡ�����ʮ����ڿ���־���ܱࡢ�����ࡢ�����ί����ί����Լ׫���ˣ��ڹ�����50����ڿ�����־���鱨��������150��ƪ������10��ƪ������������Ľ����𱭽�;������������ҽ���Ʋ����Ը��׵��о���ʵ������������4������Լ��ί7�����α�13������á�2000ǧ����ҽ����������ʱ���ܳ��������������ܳ�������ʿ���������͹��ܳ���������й�����ҽѧר�ҡ��ȶ�������ƺ�;���Ӧ����̨�塢�����ȵؽ��н�ѧ��ҽѧ���ʣ�2013��4��������̩���μӹ��ʴ�ͳҽѧ��������ᣬ��̩���׻����нӼ���', 'title': '����ҽʦ ���� ˶ʿ����ʦ �ڶ������Ҽ�����ҽҩר��ѧ������̳й���ָ����ʦ ȫ��������ҽҩר�� ҽԺ�׽���ҽ', 'imgUrl': '1111'}]

def generate_excel(rec_data):
    workbook = xlsxwriter.Workbook('C://Users//Hey//Desktop//emp.xlsx')
    worksheet = workbook.add_worksheet()

    # �趨��ʽ���Ⱥ���߸�ʽ�����Զ��壬�ֵ��и�ʽΪָ��ѡ��
     # bold���Ӵ֣�num_format:���ָ�ʽ
    bold_format = workbook.add_format({'bold': True})
    money_format = workbook.add_format({'num_format': '$#,##0'})
    date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})
    # �÷��ű��λ�ã����磺A��1��
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