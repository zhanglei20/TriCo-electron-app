#!/usr/local/Cellar/python/3.7.4_1/bin/python3 python3
# -*- coding:utf-8 -*-

import sys, os, datetime, openpyxl, json, threading
from multiprocessing import Pool

data_dict = {
    '0': {  # 酒店
        'hotelType': {
            '经济型': 1,
            '主题': 2,
            '商务型': 3,
            '公寓': 4,
            '客栈': 5,
            '青年旅社': 6,
            '度假酒店': 7,
            '星级酒店': 8
        },
        'type': {
            # '直营': 1,
            # '加盟': 2,
            # 'EGM': 4,
            '千屿2.0': 5
        },
        'brandType': {
            '千屿': 10,
            '千寻': 11
        }
    },
    '1': {  # 联系人
        'type': {  # 联系人类型
            '业主': 'owner',
            '经理': 'manager',
            '前台': 'receptionist',
            '店长': 'hotel_manager',
            '股东': 'shareholder',
            '其他': 'others'
        }
    },
    '2': {  # OTA账号
        'type': {  # ota账号类型
            '美团': 'meituan',
            '携程': 'ctrip',
            '去哪儿': 'qunar',
            '艺龙': 'elong',
            '飞猪': 'fliggy',
            '阿里商旅': 'AliTmc'
        }
    },
    '3': {  # 企业法人
        'type': {  # 企业类型
            '有限责任公司（自然人独资）': 'company',
            '个体商户': 'single',
            '个体工商户': 'single',
            '有限责任公司': 'limitedLiabilityCompany',
            '个人独资企业': 'soleProprietorship',
            '有限合伙': 'limitedPartnership',
            '股份有限公司': 'incorporatedCompany',
            '其他': 'others'
        },
        'certificateType': {  # 证件类型
            '身份证': 'CERTIFICATE_01'
        },
        'accountType': {  # 收款账号类型
            '对公账户': 1,
            '对公': 1,
            '对私账户': 2,
            '对私': 2,
            '私人账户': 2
        }
    },
    '4': {  # 房间
        'roomTypeId': {  # 房型
            '标准大床房': 20,
            '豪华大床房': 26,
            '标准双床房': 29,
            '豪华双床房': 30,
            '三人房': 33,
            '主题房': 34,
            '特惠房': 51
        },
        'bedInfoId': {
            '单人床1.0米': 1,
            '单人床1.2米': 2,
            '单人床1.35米': 3,
            '单人床1.5米': 4,
            '大床1.5米': 5,
            '大床1.8米': 6,
            '大床2.0米': 7,
            '大床2.2米': 8,
            '圆床2.0米': 9,
            '圆床2.2米': 10,
            '圆床2.4米': 11,
            '圆床2.6米': 12,
            '圆床2.8米': 13,
            '方形水床1.5米': 14,
            '方形水床1.8米': 15,
            '方形水床2.0米': 16,
            '方形水床2.2米': 17,
            '原型水床2.0米': 18,
            '原型水床2.2米': 19,
            '原型水床2.4米': 20,
            '原型水床2.6米': 21,
            '原型水床2.8米': 22
        }
    },
    '5': {  # 基础设施
        'a': 1003,  # 停车场-收费
        'b': 1004,  # 停车场-不收费
        'c': 1063,  # 无停车场
        'd': 176,  # 会议室
        'e': 3,  # WIFI覆盖
        'f': 157,  # 餐厅
        'g': 2029,  # 休息区
        'h': 180,  # 免费洗漱用品
        'i': 2024,  # 提供洗漱用品
        'k': 10031,  # 免费早餐
        'm': 10032,  # 收费早餐
        'n': 10033,  # 无早餐
        'o': 166,  # 接送服务
        'p': 170,  # 行李寄存
        'q': 171,  # 叫醒服务
        'r': 104,  # 洗衣服务
        's': 160,  # 银行卡支付
        't': 70007,  # 支付宝
        'u': 70008,  # 微信
        'v': 70009,  # 现金
    },
    '6': {
        # 'accountType': {
        #     '对公': 1,
        #     '对公账户': 1,
        #     '对私账户': 2,
        #     '对私': 2,
        #     '私人账户': 2
        # },
        'idCardType': {  # 证件类型
            '身份证': 'CERTIFICATE_01'
        }
    },
    'env': {
        'local': '127.0.0.1:5000',
        'dev': 'http://backend-product-service.dev.ahotels.tech',
        'test': 'http://backend-product-service.test.ahotels.tech',
        'uat': 'http://ali.uat.ahotels.tech/product-service',
        'prod': 'http://backend-product-service.ahotels.tech'
    }
}

# excel目录
dir = sys.argv[1]
# 环境
env = sys.argv[2]


# 读sheet
def get_sheet_data(sheet_data, index=None):
    # 全部数据
    row_list = []

    # 行号
    row_num = -1

    # json的key值
    row_header = None

    for row in sheet_data.rows:
        row_num = row_num + 1

        # 字段名
        if row_num == 0:
            row_header = [col.value for col in row]
            continue

        # 跳过表头
        if row_num == 1:
            continue

        # 一行数据
        row_data = {}

        line_data = [col.value for col in row]
        for i, value in enumerate(line_data):
            key = row_header[i]

            # 跳过key或者value为空的列
            if key is None or value is None:
                continue

            if index is not None:
                value_dict = data_dict.get(str(index), {}).get(key)
                if value_dict is not None:
                    if isinstance(value_dict, dict):
                        temp_value = value_dict.get(value)
                        if value is None:
                            print('不支持的字典: sheet=%s, key=%s, value=%s' % (index, key, temp_value))
                        value = temp_value
                    else:  # 这里是用来兼容基础设施的，基础设施的结构和其他的不一样
                        value = value_dict

            # 格式化日期
            if isinstance(value, datetime.datetime):
                value = value.strftime("%Y-%m-%d")

            row_data[key] = value

        if len(row_data.keys()) > 0:
            row_list.append(row_data)

    return row_list


# 处理协议
def get_hotel_agreement(agreement):
    not_null_list = ['totalTransCost', 'roomOyoTransCost', 'prepaymentRadio']

    for not_null_key in not_null_list:
        if not_null_key not in agreement or agreement.get(not_null_key) == '':
            print('协议字段[%s]不能为空' % not_null_key)

    for index, key in enumerate(agreement):
        if agreement.get(key) == '非必填':
            agreement[key] = ''

        if key.endswith('Url') and agreement.get(key) == '':
            agreement[key] = 'http://a'

    urls = ['businessLicensePicUrl', 'frontIdCardPicUrl', 'bankCardUrl', 'partyAfrontIdCardPicUrl',
            'reverseIdCardPicUrl']
    for url in urls:
        if url not in agreement:
            agreement[url] = 'http://a'

    # 协议类型
    agreement['projectType'] = 'ISLANDS_2'
    agreement['commissionDateTypeName'] = '控价日'

    mdm_audit_time_key = 'mdmAuditTime'
    if mdm_audit_time_key not in agreement or agreement[mdm_audit_time_key] is None:
        # 当前时间+8小时，因为需要在submit时间之后
        mdm_audit_time = datetime.datetime.now() + datetime.timedelta(hours=8)
        agreement[mdm_audit_time_key] = mdm_audit_time.strftime('%Y-%m-%d %H:%M:%S')

    prepayment_radio = agreement['prepaymentRadio']
    if prepayment_radio < 1:
        prepayment_radio = prepayment_radio * 100
    agreement['prepaymentRadio'] = prepayment_radio

    agreement['guaranteeIncomes'] = [{
        "month": "1",
        "money": agreement.pop('janAmg'),
        "vmgMoney": agreement.pop('janVmg')
    }, {
        "month": "2",
        "money": agreement.pop('febAmg'),
        "vmgMoney": agreement.pop('febVmg')
    }, {
        "month": "3",
        "money": agreement.pop('marchAmg'),
        "vmgMoney": agreement.pop('marchVmg')
    }, {
        "month": "4",
        "money": agreement.pop('aprilAmg'),
        "vmgMoney": agreement.pop('aprilVmg')
    }, {
        "month": "5",
        "money": agreement.pop('mayAmg'),
        "vmgMoney": agreement.pop('mayVmg')
    }, {
        "month": "6",
        "money": agreement.pop('juneAmg'),
        "vmgMoney": agreement.pop('juneVmg')
    }, {
        "month": "7",
        "money": agreement.pop('julyAmg'),
        "vmgMoney": agreement.pop('julyVmg')
    }, {
        "month": "8",
        "money": agreement.pop('augustAmg'),
        "vmgMoney": agreement.pop('augustVmg')
    }, {
        "month": "9",
        "money": agreement.pop('septAmg'),
        "vmgMoney": agreement.pop('septVmg')
    }, {
        "month": "10",
        "money": agreement.pop('octAmg'),
        "vmgMoney": agreement.pop('octVmg')
    }, {
        "month": "11",
        "money": agreement.pop('noveAmg'),
        "vmgMoney": agreement.pop('noveVmg')
    }, {
        "month": "12",
        "money": agreement.pop('deceAmg'),
        "vmgMoney": agreement.pop('deceVmg')
    }]

    return agreement


# 二次处理房间信息
# 从sheet中读取到的是根据楼层分组的房间，这里把房间号和楼层进行交集处理
def get_rooms(floor_rooms, hotel_name):
    ret_rooms = {}

    room_num_check_list = []
    room_sum = 0
    print('<br>')
    print(hotel_name)

    for index, room_type_floor in enumerate(floor_rooms):
        room_nums = room_type_floor.get('roomNo')

        if room_nums is None:
            continue

        if isinstance(room_nums, int):
            room_nums = str(room_nums)
        row_index = index + 3
        room_nums = room_nums.replace('、', ',').replace('，', ',').replace('.', ',')

        room_type_id = room_type_floor.get('roomTypeId')
        if room_type_id is None:
            print('<p style="color:red">' + hotel_name + ',第%s行,房型错误</p>' % index + 3)

        ret_room_type = ret_rooms.get(room_type_id, {
            'roomTypeId': room_type_id,
            'weekdayPrice': room_type_floor.get('weekdayPrice'),
            'weekendPrice': room_type_floor.get('weekendPrice'),
            'roomList': []
        })

        floor = room_type_floor.get('floor')
        room_size = room_type_floor.get('size')
        bed_count = room_type_floor.get('count')
        bed_info_id = room_type_floor.get('bedInfoId')
        print('房型: %s, 楼层: %s, 房间号: %s, 面积: %s, 床数: %s, 床型: %s'
              % (room_type_id, floor, room_nums, room_size, bed_count, bed_info_id))
        print('<br>')

        if floor is None:
            print('<p style="color:red">' + hotel_name + '第%s行楼层缺失</p>' % row_index)
        if room_size is None:
            print('<p style="color:red">' + hotel_name + '第%s行房间面积缺失</p>' % row_index)
        if bed_count is None:
            print('<p style="color:red">' + hotel_name + '第%s行床数缺失</p>' % row_index)
        if bed_info_id is None:
            print('<p style="color:red">' + hotel_name + '第%s行床型错误</p>' % row_index)
        room_list = []
        check = ',' not in room_nums
        room_nums_arr = room_nums.split(',')
        for room_i, room_no in enumerate(room_nums_arr):
            if room_no is None or room_no == '':
                continue

            # 去掉前后空格
            room_no = room_no.strip()

            if room_no in room_num_check_list:
                print('<p style="color:red">' + hotel_name + ',房间号%s已存在</p>' % room_no)
            # 长度大于6位 且 并未包含分隔符
            if len(room_no) >= 6 and check:
                print('<p style="color:red">' + hotel_name + ',房间号%s格式不对</p>' % room_no)

            room_num_check_list.append(room_no)

            room_list.append({
                'floor': floor,
                'roomNo': room_no,
                'size': room_size,
                'status': 1,
                'bedInfoList': [{
                    'count': bed_count,
                    'bedInfoId': bed_info_id
                }]
            })

        if len(room_list) > 0:
            ret_room_type.get('roomList').extend(room_list)
        # ret_rooms[room_type_id] = ret_room_type
        room_sum = room_sum + len(ret_room_type.get('roomList'))
    return room_sum


excel_sheets = {
    '酒店': 0,
    '酒店(勿动)': 0,
    '酒店（勿动）': 0,
    '联系人': 1,
    '联系人(勿动)': 1,
    '联系人（勿动）': 1,
    'OTA账号': 2,
    'OTA账号(勿动)': 2,
    'OTA账号（勿动）': 2,
    '企业法人': 3,
    '企业法人(勿动)': 3,
    '企业法人（勿动）': 3,
    '房间': 4,
    '房型房间信息(编辑区)': 4,
    '房型房间信息（编辑区）': 4,
    '房型房间信息（填写区）': 4,
    '酒店设施': 5,
    '酒店设施(勿动)': 5,
    '酒店设施（勿动）': 5
    # '协议信息': 6,
    # '协议信息（勿动）': 6
}


# 读excel
def read_xlsx(file_path, file_name):
    ret_json_hotel = {
        "type": "qianyu",
        "operator": "qy_import"
    }

    wb = openpyxl.load_workbook(file_path, data_only=True)
    for name, index in excel_sheets.items():
        if name in wb.sheetnames:
            sheet_index = excel_sheets.get(name)
            sheet_data = get_sheet_data(wb.get_sheet_by_name(name), sheet_index)

            if len(sheet_data) == 0:
                continue

            # 1.酒店信息
            if sheet_index == 0:
                hotel_info = sheet_data[0]  # 读取酒店信息
                if str(hotel_info['stateId']) is '#REF!':
                    print('<p style="color:red">' + file_name + ',酒店 省 填写错误</p>')
                if str(hotel_info['cityId']) is '#REF!':
                    print('<p style="color:red">' + file_name + ',酒店 市 填写错误</p>')
                if str(hotel_info['clusterId']) is '#REF!':
                    print('<p style="color:red">' + file_name + ',酒店 区 填写错误</p>')
                if str(hotel_info['streetId']) is '#REF!':
                    print('<p style="color:red">' + file_name + ',酒店 街道 填写错误</p>')
                sign_room_num = hotel_info.get('signRoomNum')

            # 2.联系人
            if sheet_index == 1:
                contacts = []
                for index, contact in enumerate(sheet_data):
                    email = contact.get('email')
                    if email == '' or email == '0' or email == 0:
                        contact.pop('email')

                    if contact.get('name') is not None:
                        contacts.append(contact)

            # 3.OTA账号
            if sheet_index == 2:
                for index, ota_account in enumerate(sheet_data):
                    if ota_account.get('name') is '#N/A' or ota_account.get('password') is '#N/A' or ota_account.get(
                            'name') is '#N/A' or ota_account.get('password') is '#N/A':
                        print('<p style="color:red">' + file_name + ',OTA账号存在脏数据</p>')

            # 4.企业法人和银行信息
            if sheet_index == 3:
                legal_bank_info = sheet_data[0]

                # ret_json_hotel['legalEntities'] = [{
                #     'name': legal_bank_info.get('name'),
                #     'type': legal_bank_info.get('type'),
                #     'termOfOperation': legal_bank_info.get('termOfOperation'),
                #     'licenseNo': legal_bank_info.get('licenseNo'),
                #     'certificateType': legal_bank_info.get('certificateType'),
                #     'certificateNo': legal_bank_info.get('certificateNo')
                # }]

                if legal_bank_info.get('type') is None:
                    print('<p style="color:red">' + file_name + ',企业类型有误</p>')

                if legal_bank_info.get('accountType') is None:
                    print('<p style="color:red">' + file_name + ',账号类型有误</p>')

                # bank_interbank_number = legal_bank_info.get('bankInterbankNumber')
                # if bank_interbank_number == '' or bank_interbank_number is None:
                #     bank_interbank_number = 'a'
                #
                # bank_ddress = legal_bank_info.get('bankAddress')
                # if bank_ddress == '' or bank_ddress is None:
                #     bank_ddress = 'a'
                #
                # ret_json_hotel['bankAccountInfos'] = [{
                #     'receivingParty': legal_bank_info.get('receivingParty'),
                #     'contactTelephone': legal_bank_info.get('contactTelephone'),
                #     'accountType': legal_bank_info.get('accountType'),
                #     'bankAccount': legal_bank_info.get('bankAccount'),
                #     'openingBank': legal_bank_info.get('openingBank'),
                #     'branchOpeningBank': legal_bank_info.get('branchOpeningBank'),
                #     'bankInterbankNumber': bank_interbank_number,
                #     'bankAddress': bank_ddress
                # }]

            # 5.房间
            if sheet_index == 4:
                # ret_json_hotel['room'] = get_rooms(sheet_data)
                real_room_num = get_rooms(sheet_data, file_name)
                if sign_room_num - real_room_num:
                    print('请检查房间数量，签约房间数为:%s,当前房间数为:%s' % (sign_room_num, real_room_num))

            # 6.酒店设施
            if sheet_index == 5:
                amenity_list = []
                if sheet_data is not None and len(sheet_data) > 0:
                    for key, amenity_id in sheet_data[0].items():
                        if isinstance(amenity_id, int):
                            amenity_list.append(amenity_id)

                ret_json_hotel['amenityList'] = amenity_list

    return ret_json_hotel


# 递归目录  考虑多线程操作 减少耗时
def list_dir(path):
    now = datetime.datetime.now()  # 开始计时
    print('开始时间：' + now.strftime("%Y-%m-%d %H:%M:%S"))
    print('<br>')

    file_names = os.listdir(path)
    file_names.sort()
    # 线程组
    # thread = []
    # 进程组
    # process = []
    process = Pool(20)
    # res_data = []

    for i in range(len(file_names)):
        if not file_names[i].startswith('~$') and file_names[i].endswith('.xlsx'):
            # print('<br>')
            # print('<p>%s: %s</p>' % (i, file_names[i]))
            # print('<br>')
            file_name = path + "/" + file_names[i]  # 要获取的excel地址
            # t = threading.Thread(target=read_xlsx, args=(file_name,))
            # thread.append(t)
            # p = Process(target=read_xlsx, args=(file_name,))
            # process.append(p)
            process.apply_async(read_xlsx, (file_name, file_names[i]))
            # res_data.append(res_d)

    # 多线程模式 不适用 GIL的存在，使得Python在同一时间只能运行一个线程，所以只占用了一个CPU
    # thread_num = len(thread)
    # print(thread_num)
    # for i in range(len(thread)):
    #     thread[i].start()
    #
    # for i in range(len(thread)):
    #     thread[i].join()

    # 多进程模式 如果进程都需要写入同一个文件，那么就会出现多个进程争用资源
    # process_num = len(process)
    # for i in range(process_num):
    #     process[i].start()
    #
    # for i in range(process_num):
    #     process[i].join()

    process.close()
    process.join()

    end = datetime.datetime.now()  # 结束计时
    print('<br>')
    print('结束时间：' + end.strftime("%Y-%m-%d %H:%M:%S"))
    print('<br>')
    print('程序耗时： ' + str(end - now))


list_dir(dir)
