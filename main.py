__author__ = 'Administrator'
#coding:cp936
import xlrd
import xlwt
SRC_FILE_NAME = u"����.xlsx"
DST_FILE_NAME = u"result.xls"
EXIST_FLAG = 0
NOT_EXIST_FLAG = 1
LINE_COUNT = 53
SHEET_HEADER = (
    u"��������",
    u"������",
    u"����",
    u"�ջ���",
    u"�ջ��ֻ�",
    u"��ַ",
    u"�ʱ�",
    u"�ļ�������",
    u"�ļ��˵�ַ",
    u"�ļ��˵绰",
    u"�ļ����ʱ�",
)

template_file_header = (
    u"�ļ�������",
    u"�ļ�����ϵ��ʽ",
    u"�ļ��˵�ַ",
    u"�ռ�������",
    u"�ռ�����ϵ��ʽ",
    u"�ռ��˵�ַ",
    u"�������",
    u"�ռ�����ϵ��ʽ��2��",
    u"�ռ����ʱ�",
    u"�ռ��˹�˾",
    u"����ʡ/ֱϽ��",
    u"��������",
    u"������/��",
    u"��Ʒ����",
    u"��Ʒ����",
    u"��ʱ��",
    u"��ע",
    u"ҵ������",
    u"�ڼ���Ϣ",
    u"����һ",
    u"���׶�",
    u"�ļ����ʱ�",
    u"�ռ���Ӧ������",
    u"�ռ���Ӧ�����ʣ���д��",
    u"���۽��",
    u"�ļ��˹�˾",
)
def main(filename):
    result_list = []
    try:
        work_book = xlrd.open_workbook(SRC_FILE_NAME)
        data_sheet = work_book.sheet_by_name(u'Sheet1')
    except:
        print "���ļ�ʧ�ܣ��������ļ��Ѿ����򿪻��߲�����"
        return False
    else:
        nrows = data_sheet.nrows
        for i in range(1, nrows):
            order_info = data_sheet.row_values(i)
            ret = find_order_code(order_info, result_list)
            if EXIST_FLAG == ret:
                continue
            elif NOT_EXIST_FLAG == ret:
                new_item = {}
                new_item["order_code"] = order_info[0]
                new_item["bar_code"] = []
                new_item["bar_code"].append(
                    {
                        "bar_code":order_info[1],
                        "order_quantity":order_info[2],
                        "order_weight":order_info[3]
                    }
                )
                new_item["consignee_name"] = order_info[4]
                new_item["consignee_phone_number"] = order_info[5]
                new_item["consignee_address"] = order_info[6]
                new_item["consignee_postcode"] = order_info[7]
                new_item["sender_name"] = order_info[8]
                new_item["sender_address"] = order_info[9]
                new_item["sender_phone_number"] = order_info[10]
                new_item["sender_postcode"] = order_info[11]
                result_list.append(new_item)
        wbk = xlwt.Workbook()
        sheet = wbk.add_sheet(u"Sheet1")
        write_result_sheet_to_template_file(result_list, sheet)
        import time
        dst_file_name = time.strftime('result-%Y-%m-%d-%H-%M-%S.xls', time.localtime(time.time()))
        wbk.save(dst_file_name)
        print "���ɽ��������%s��" % dst_file_name
        return True

def find_order_code(order_info, result_list):
    order_code = order_info[0]
    bar_code = order_info[1]
    order_quantity = order_info[2]
    order_weight = order_info[3]
    consignee_name = order_info[4]
    consignee_phone_number = order_info[5]
    consignee_address = order_info[6]
    consignee_postcode = order_info[7]
    sender_name = order_info[8]
    sender_address = order_info[9]
    sender_phone_number = order_info[10]
    sender_postcode = order_info[11]
    for i in result_list:
        if i["order_code"] == order_code:
            i["bar_code"].append(
                {
                    "bar_code":bar_code,
                    "order_quantity":order_quantity,
                    "order_weight":order_weight
                }
            )
            return EXIST_FLAG
        else:
            continue
    return NOT_EXIST_FLAG
def write_sheet_header(data_sheet, sheet_header):
    for i in range(0, len(sheet_header)):
        data_sheet.write(0, i, sheet_header[i])

def write_result_sheet(result_list, sheet):
    if isinstance(result_list, list):
        nrows = len(result_list)
        write_sheet_header(sheet, SHEET_HEADER)
        for i in range(1, nrows + 1):
            sheet.write(i, 0, result_list[i-1]["order_code"])
            #t_bc��������ɵĴ�ӡ�������ַ���
            t_bc = ""
            #��������
            total_weight = 0.0
            for bc in result_list[i-1]["bar_code"]:
                total_weight += float(bc["order_quantity"]) * float(bc["order_weight"])
            #׼�������������ַ���
            total_barcode_num = len(result_list[i-1]["bar_code"])
            for x in range(0, total_barcode_num):
                s_bc_str = "%s-%d" % (
                    result_list[i-1]["bar_code"][x]["bar_code"],
                    result_list[i-1]["bar_code"][x]["order_quantity"]
                )
                s_bc_str += (26 - len(s_bc_str)) * " "
                if x%2 != 0 and x < (total_barcode_num - 1):
                    s_bc_str += '\n'
                t_bc += s_bc_str
            sheet.write(i, 1, t_bc)
            sheet.write(i, 2, str(total_weight))
            sheet.write(i, 3, result_list[i-1]["consignee_name"])
            sheet.write(i, 4, result_list[i-1]["consignee_phone_number"])
            sheet.write(i, 5, result_list[i-1]["consignee_address"])
            sheet.write(i, 6, result_list[i-1]["consignee_postcode"])
            sheet.write(i, 7, result_list[i-1]["sender_name"])
            sheet.write(i, 8, result_list[i-1]["sender_address"])
            sheet.write(i, 9, result_list[i-1]["sender_phone_number"])
            sheet.write(i, 10, result_list[i-1]["sender_postcode"])
        return True
    else:
        return False

def write_result_sheet_to_template_file(result_list, sheet):
    if isinstance(result_list, list):
        nrows = len(result_list)
        write_sheet_header(sheet, template_file_header)
        for i in range(1, nrows + 1):
            sheet.write(i, 6, result_list[i-1]["order_code"])
            #t_bc��������ɵĴ�ӡ�������ַ���
            t_bc = ""
            #��������
            total_weight = 0.0
            for bc in result_list[i-1]["bar_code"]:
                total_weight += float(bc["order_quantity"]) * float(bc["order_weight"])
                #׼�������������ַ���
            total_barcode_num = len(result_list[i-1]["bar_code"])
            for x in range(0, total_barcode_num):
                s_bc_str = "%s-%d" % (
                    result_list[i-1]["bar_code"][x]["bar_code"],
                    result_list[i-1]["bar_code"][x]["order_quantity"]
                )
                s_bc_str += (26 - len(s_bc_str)) * " "
                if x%2 != 0 and x < (total_barcode_num - 1):
                    s_bc_str += '\n'
                t_bc += s_bc_str
            sheet.write(i, 18, t_bc)
            sheet.write(i, 13, str(total_weight))
            sheet.write(i, 3, result_list[i-1]["consignee_name"])
            sheet.write(i, 4, result_list[i-1]["consignee_phone_number"])
            sheet.write(i, 5, result_list[i-1]["consignee_address"])
            sheet.write(i, 8, result_list[i-1]["consignee_postcode"])
            sheet.write(i, 0, result_list[i-1]["sender_name"])
            sheet.write(i, 2, result_list[i-1]["sender_address"])
            sheet.write(i, 1, result_list[i-1]["sender_phone_number"])
            sheet.write(i, 21, result_list[i-1]["sender_postcode"])
        return True
    else:
        return False
if __name__ == '__main__':
    while True:
        filename = raw_input("�������ļ���:")
        main(filename)