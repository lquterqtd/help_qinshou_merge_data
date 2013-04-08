__author__ = 'Administrator'
#coding:cp936
import xlrd
import xlwt
SRC_FILE_NAME = u"数据.xlsx"
DST_FILE_NAME = u"result.xls"
EXIST_FLAG = 0
NOT_EXIST_FLAG = 1
LINE_COUNT = 53
SHEET_HEADER = (
    u"订单编码",
    u"条形码",
    u"重量",
    u"收货人",
    u"收货手机",
    u"地址",
    u"邮编",
    u"寄件人姓名",
    u"寄件人地址",
    u"寄件人电话",
    u"寄件人邮编",
)
def main(filename):
    result_list = []
    try:
        work_book = xlrd.open_workbook(SRC_FILE_NAME)
        data_sheet = work_book.sheet_by_name(u'Sheet1')
    except:
        print "打开文件失败，可能是文件已经被打开或者不存在"
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
        write_result_sheet(result_list, sheet)
        import time
        dst_file_name = time.strftime('result-%Y-%m-%d-%H-%M-%S.xls', time.localtime(time.time()))
        wbk.save(dst_file_name)
        print "生成结果保存在%s中" % dst_file_name
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
def write_sheet_header(data_sheet):
    for i in range(0, len(SHEET_HEADER)):
        data_sheet.write(0, i, SHEET_HEADER[i])

def write_result_sheet(result_list, sheet):
    if isinstance(result_list, list):
        nrows = len(result_list)
        write_sheet_header(sheet)
        for i in range(1, nrows + 1):
            sheet.write(i, 0, result_list[i-1]["order_code"])
            #t_bc是最后生成的打印条形码字符串
            t_bc = ""
            #求总重量
            total_weight = 0.0
            for bc in result_list[i-1]["bar_code"]:
                total_weight += float(bc["order_quantity"]) * float(bc["order_weight"])
            #准备生成条形码字符串
            total_barcode_num = len(result_list[i-1]["bar_code"])
            for x in range(0, total_barcode_num):
                if x < (total_barcode_num - 1) and x%2 != 0:
                    #不是最后一个条形码但是一行的结束
                    s_bc_str = "%s-%d\n" % (
                        result_list[i-1]["bar_code"][x]["bar_code"],
                        result_list[i-1]["bar_code"][x]["order_quantity"]
                    )
                    space_num = 0
                elif x < (total_barcode_num - 1) and  x%2 == 0:
                    #不是最后一个条形码也不是一行的结束
                    s_bc_str = "%s,%d" % (
                        result_list[i-1]["bar_code"][x]["bar_code"],
                        result_list[i-1]["bar_code"][x]["order_quantity"]
                    )
                    n_bc_str = "%s,%d\n" % (
                        result_list[i-1]["bar_code"][x+1]["bar_code"],
                        result_list[i-1]["bar_code"][x+1]["order_quantity"]
                    )
                    space_num = LINE_COUNT - len(s_bc_str) - len(n_bc_str)
                elif x >= (total_barcode_num - 1):
                    #是最后一个条形码且是一行的结束
                    #是最后一个条形码但不是一行的结束
                    s_bc_str = "%s,%d" % (
                        result_list[i-1]["bar_code"][x]["bar_code"],
                        result_list[i-1]["bar_code"][x]["order_quantity"]
                    )
                    space_num = 0
                t_bc += s_bc_str
                t_bc += (space_num * ' ')
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
if __name__ == '__main__':
    while True:
        filename = raw_input("请输入文件名:")
        main(filename)