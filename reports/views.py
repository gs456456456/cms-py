# -*- encoding:utf-8 -*-
import json
import datetime
import hashlib
import arrow
import requests
import xlsxwriter
from django.shortcuts import render
from django.http import HttpResponse
from .forms import ReportForm
from .models import Machine, Tag, Config
import time
# from django.http import StreamingHttpResponse, HttpResponseRedirect


# Create your views here.


def index(request):
    # return HttpResponse("Hello, world. You're at the polls index.")
    form = ReportForm()
    context = {'form': form}
    return render(request, 'reports/index.html', context)


def vote(request):
    # return HttpResponse("Hello, world. You're voting from index.")
    timestart = time.time()
    report_type = request.POST['报表类型']
    report_date = request.POST['日期']
    machine = Machine.objects.all().values_list('machine_id')
    tag = Tag.objects.all().values_list('tag_id')
    # s = "Report Type: %s; Report Date: %s" % (report_type, report_date)
    # 从云端获取数据
    if report_type == 'daily':
        date_begin_utc = arrow.get(report_date)
        # 默认时区为0,先减8小时,随后修改时区,匹配本地时间
        date_begin = date_begin_utc.replace(hours=-8)
        date_end = date_begin.replace(days=+1)
        # 请求数据的起始时间
        startTime = date_begin.to('local')
        endTime = date_end.to('local')
        daily = '日报'.encode('gb2312')
        filename = ''.join('{0} {1}'.format('日报',str(report_date)))
        # print(filename)
    elif report_type == 'weekly':
        report_datetime = datetime.date(int(report_date.split(
            '-')[0]), int(report_date.split('-')[1]), int(report_date.split('-')[2]))
        report_datetime_week = int(report_datetime.weekday())
        date_now = arrow.get(report_date)
        date_begin_utc = date_now.replace(days=-report_datetime_week)
        date_begin = date_begin_utc.replace(hours=-8)
        date_end = date_begin.replace(days=+7)
        startTime = date_begin.to('local')
        endTime = date_end.to('local')
        filename = ''.join('{0} {1}'.format('周报',str(report_date)))
        # print(filename)
    elif report_type == 'monthly':
        report_date_first = report_date[:-2] + '01'
        date_begin_utc = arrow.get(report_date_first)
        date_begin = date_begin_utc.replace(hours=-8)
        date_end = date_begin.replace(months=+1)
        startTime = date_begin.to('local')
        endTime = date_end.to('local')
        print(startTime)
        print(endTime)
        filename = ''.join('{0} {1}'.format('月报',str(report_date[:-3])))
        # print(filename)
    t0 = arrow.utcnow().float_timestamp
    # 网关的ID，每个场地都不一样
    # 获取方法见 08_定制开发看板使用说明_BEDRAGON_V1.0.pdf 3.2.2节
    # machine_id = "yaYFRz.255.1"
    machine_ids = []
    for item in machine:
        for x in item:
            machine_ids.append(x)
    print(machine_ids)
    tag_ids = []
    for item in tag:
        for x in item:
            tag_ids.append(x)
    print(tag_ids)
    # 获取机器种类
    # 在网关中定义的变量ID，有多少个id就对应多少个料仓
    # 以50开始，每个料仓递增2
    # tags = [50, 52, 54, 56, 58, 60, 62, 64, 66, 68, 70, 72]
    # tags = [21, 22, 23, 24, 25, 26, 27, 28, 29,30]

    # 从云端获取数据
    r0 = getData(startTime, endTime, machine_ids, tag_ids)
    # for debug only
    write_json_file("result_r0.json", r0)

    if r0 is None or 'error' in r0:
        return HttpResponse("Network Failure!")

    # 转换时间戳格式
    r1 = convertDateTime(r0)
    # print(r1)

    # for debug only
    write_json_file("result_r1.json", r1)

    t1 = arrow.utcnow().float_timestamp
    delta_t = t1 - t0
    print("Time cost: %.3f seconds" % delta_t)

    # excel 相关操作
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename=%s.xlsx' % filename.encode('utf-8').decode('ISO-8859-1')
    row = 0
    col = 0
    colname = ['设备编号', '员工编号', '换班时间', '规格', '克重', '报警数', '损耗率%', '客户编码','产量']
    workbook = xlsxwriter.Workbook(response, {'in_memory': True})
    format = workbook.add_format()
    worksheet = workbook.add_worksheet()
    format.set_font_color('red')
    # 新建列名
    for item in colname:
        worksheet.write(row, col, item)
        col += 1
    # 数据填充
    row = 1
    col = -1
    row2 = 1
    col2 = 7
    del_list = del_repeat_list()
    for x in [1,2,3,4,5,6,8,9]:
        datavalue = getJsonData(x)
        if datavalue:
            z = 0
            for y in del_list:
                datavalue.pop(y - z)
                z += 1
            col += 1
            row = 1
        for item in datavalue:
            if int(x)<7:
                worksheet.write(row, col, item)
                row += 1
            elif int(x)>7:
                worksheet.write(row2, col2, item)
                row2 += 1
    #损耗率超过2红字
    row = 1
    col = 5
    lost = del_repeat_list()
    ldatavalue = getJsonData(7)
    if ldatavalue:
        z = 0
        for y in del_list:
            ldatavalue.pop(y - z)
            z += 1
        col += 1
        row = 1
    for item in ldatavalue:
        if float(item)>2:
            worksheet.write(row, col, item,format)
        else:
            worksheet.write(row, col, item)
        row += 1


    # 将shiftTime时间戳列转化为日期格式
    row = 1
    col = 2
    datavalueshifttime = getJsonData(3)
    z = 0
    for y in del_list:
        datavalueshifttime.pop(y - z)
        z += 1
    shifttime = []
    for item in datavalueshifttime:
        if item == 'None\n':
            shifttime.append('None')
        else:
            localtime = arrow.get(item).to('local')
            shifttime.append(localtime.format())
    for item in shifttime:
        item = item[2:-6]
        worksheet.write(row, col, item)
        row += 1

    #求出产量的差值
    row = 1
    col = 8
    datavalueproductcount = getJsonData(9)
    z = 0
    if datavalueproductcount:
        for y in del_list:
            datavalueproductcount.pop(y - z)
            z += 1
        productcount = []
        for num,item in enumerate(datavalueproductcount):
            if item == 'None\n' or item == '0\n':
                datavalueproductcount[num] = '0\n'
                productcount.append('0\n')
            else:
                if num>0:
                    if int(datavalueproductcount[num-1][:-1])!= 0:
                        productcount.append(int(datavalueproductcount[num][:-1])-(int(datavalueproductcount[num-1][:-1])))
                    else:
                        productcount.append(int(datavalueproductcount[num][:-1])-int(datavalueproductcount[num-2][:-1]))
        for item in productcount:
            if item == '0\n':
                item = 'None'
            worksheet.write(row, col, item)
            row += 1
        # print(productcount)
        productcount.append('...')
    workbook.close()
    timeover = time.time()
    print(timeover-timestart)
    return response

def del_repeat():
    #删除flag重复项以及所对应的其他参数
    pass


def convertDateTime(r):
    if r is None:
        return []
    for item in r:
        value_list = item['TargetEnergyData'][0]['EnergyData']
        for value in value_list:
            t0 = value['UtcTime']
            t1 = energyDateTime2DTString(t0, '+08:00', 'YYYY-MM-DD HH:mm:ss')
            value['LocalTime'] = t1
            t2 = energyDateTime2DTString(t0, '+00:00', 'YYYY-MM-DD HH:mm:ss')
            value['UtcTime'] = t2
    return r


def write_json_file(path, contentOject):
    # json_string = json.dumps(contentOject, 'utf-8')
    json_string = json.dumps(contentOject, sort_keys=True, indent=4,
                             ensure_ascii=False)
    try:
        config_file = open(path, "w")
        config_file.write(json_string)
        r = True
    except IOError as err:
        errorStr = 'File Error:' + str(err)
        print(errorStr)
        r = False
    finally:
        if 'config_file' in locals():
            config_file.close()
    return r


def energyDateTime2DTString(energyDateTime, timezone, tzformat):
    s = energyDateTime.replace("/Date(", "").replace(")/", "")
    # remove the last 3 zeros, convert from millisecond to second
    s = s[:len(s) - 3]
    # the cloud system use UTC timestamp
    try:
        s = int(s)
        adt = arrow.get(s)
    except Exception as e:
        print("energyDateTime2DTString: " + str(e))
        return ''
    else:
        r = adt.to(timezone).format(tzformat)
        return r


def getJsonData(x):
    # x为获取数据的种类
    datavalue = []
    filename = 'result_r1.json'
    with open(filename) as f:
        data = json.load(f)
        for y in data[x]['TargetEnergyData']:
            energydata = y['EnergyData']
            for z in energydata:
                # if not z['DataValue'] is None:
                datavalue.append(str(z['DataValue']))
            with open('test.txt', 'w') as b:
                datavalue = [line + '\n' for line in datavalue]
                b.writelines(datavalue)
    return datavalue


def del_repeat_list():
    # 得到删除flag重复项以及所对应的其他参数列表
    data = getJsonData(0)
    data_repeat_pos_list = []
    del_list = []
    for item in data:
        data_repeat_pos_zero = [i for i, v in enumerate(data) if v == "None\n"]
        data_repeat_pos = [i for i, v in enumerate(data) if v == item]
        data_repeat_pos_havezero = data_repeat_pos[1:]
        data_repeat_pos_list.append(data_repeat_pos_havezero)
        data_repeat_pos_list.append(data_repeat_pos_zero)
    for item in data_repeat_pos_list:
        if item not in del_list and item:
            for x in item:
                del_list.append(x)
    del_list = list(set(del_list))
    return del_list

def getData(startTime, endTime, machine_ids, tags):
    # input paramters:
    # startTime = '2017-03-08T08:00:00+08:00'
    # endTime = '2017-03-08T10:00:00+08:00'
    # machine_id = "yaYFRz.255.1"
    # for 12 silos
    # tags = [50, 52, 54, 56, 58, 60, 62, 64, 66, 68, 70, 72]
    # -------------------------------------------------------------------------
    # fixed parameters

    customerCode = Config.objects.filter(id=1).values('customer_code').first()['customer_code']
    appkey = Config.objects.filter(id=1).values('app_key').first()['app_key']
    appsecret = Config.objects.filter(id=1).values('app_secret').first()['app_secret']
    apiurl = Config.objects.filter(id=1).values('api_url').first()['api_url']
    # -------------------------------------------------------------------------
    headers = {}
    headers['Content-Type'] = 'application/json'
    headers['X-Auth-AppKey'] = appkey

    payload = {}
    payload['CustomerCode'] = customerCode
    payload['ProcessPrecision'] = True
    payload['TagCodes'] = []
    for machine_id in machine_ids:
        for tag_id in tags:
            read_tag = str(machine_id) + "." + str(tag_id)
            payload['TagCodes'].append(read_tag)
    # JHXING / BEDRAGON
    # payload['TagCodes'] = ['yaYFRz.255.1.50',
    #                        'yaYFRz.255.1.52',
    #                        'yaYFRz.255.1.54']
    # print(payload['TagCodes'])
    ViewOption = {}
    # ViewOption['DataOption'] = None
    ViewOption['DataOption'] = {"OriginalValue": True}
    ViewOption['ValueOptions'] = [1]
    ViewOption['Step'] = 0
    slot1 = {}

    # the cloud system use UTC timestamp
    # ts1 = arrow.utcnow().timestamp
    # ts0 = ts1 - 60 * 20
    # slot1['StartTime'] = "/Date(%d000)/" % ts0
    # slot1['EndTime'] = "/Date(%d000)/" % ts1

    try:
        a = arrow.get(startTime)
        b = arrow.get(endTime)
    except Exception as e:
        print(e)
        a = arrow.utcnow().timestamp
        b = arrow.utcnow().timestamp - 60 * 20
    else:
        a = a.timestamp
        b = b.timestamp
    slot1['StartTime'] = "/Date(%d000)/" % a
    slot1['EndTime'] = "/Date(%d000)/" % b

    ViewOption['TimeRanges'] = [slot1]
    payload['ViewOption'] = ViewOption

    body = json.dumps(payload)
    m = hashlib.md5()
    s = appkey + body + appsecret
    m.update(s.encode('utf-8'))
    appfig = m.hexdigest()
    # print(appkey)
    # print(body)
    # print(appsecret)
    # print(appfig)
    headers['X-Auth-Fig'] = appfig

    try:
        r = requests.post(apiurl, data=body, headers=headers, timeout=10)
    except Exception as e:
        print(e)
        return None

    if r.status_code != 200:
        info = "get data from openAPI failed"
        print(info)
        return None

    # print(r)
    # print(r.json())
    return r.json()
