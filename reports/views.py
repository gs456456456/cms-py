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
from dwebsocket import require_websocket
import queue
import re
import datetime
# from django.http import StreamingHttpResponse, HttpResponseRedirect


def index(request):
    # return HttpResponse("Hello, world. You're at the polls index.")
    form = ReportForm()
    context = {'form': form}
    return render(request, 'reports/index.html', context)

# Create your views here.
@require_websocket
def echo_once(request):
    message = request.websocket.wait()
    request.websocket.send(message)
    return render(request,'webs.html')

def myview(request):
    return render(request,'reports/newpage.html')


####################################获取云端数据

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

############json处理############

def readJson():
    filename = 'result_r1.json'
    with open(filename) as f:
        data = json.load(f)
    return data


def getOneJsonData(data,x):
    datavalue = []
    for y in data[x]['TargetEnergyData']:
        energydata = y['EnergyData']
        for z in energydata:
            datavalue.append(str(z['DataValue']))
    return datavalue


def gettotalJsonData():
    final_list = []
    num_list = []
    zip_list = []
    data = readJson()
    for i in range(12):
        final_list.append(getOneJsonData(data,i))
    for num,value in enumerate(final_list[11]):
        if value == 'None':
            num_list.append(num)
    for i in final_list:
        k = 0
        for j in num_list:
            i.pop(j-k)
            k+=1
    a = final_list[0]
    b = final_list[1]
    c = final_list[2]
    d = final_list[3]
    e = final_list[4]
    f = final_list[5]
    g = final_list[6]
    h = final_list[7]
    i = final_list[8]
    j = final_list[9]
    k = final_list[10]
    zip_list = list(zip(a,b,c,d,e,f,g,h,i,j,k))
    return zip_list


def testview(request):
    a = gettotalJsonData()
    print(a[11])

    return HttpResponse('ok')

# def del_repeat_list():
#     # 得到删除flag重复项以及所对应的其他参数列表
#
#     data = getJsonData(11)
#     data_repeat_pos_list = []
#     del_list = []
#     for item in data:
#         data_repeat_pos_zero = [i for i, v in enumerate(data) if v == "None\n"]
#         data_repeat_pos = [i for i, v in enumerate(data) if v == item]
#         data_repeat_pos_havezero = data_repeat_pos[1:]
#         data_repeat_pos_list.append(data_repeat_pos_havezero)
#         data_repeat_pos_list.append(data_repeat_pos_zero)
#     for item in data_repeat_pos_list:
#         if item not in del_list and item:
#             for x in item:
#                 del_list.append(x)
#     del_list = list(set(del_list))
#     return del_list




##################


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

############################################
###2017-9-7


def jsonSave(request,startTime,endTime):
    machine = Machine.objects.all().values_list('machine_id')
    tag = Tag.objects.all().values_list('tag_id')
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

def time_change(mytime):
    newlist = []
    for x in mytime:
        a = arrow.get(x)
        b = a.strftime('%Y-%m-%d %H:%M:%S')
        newlist.append(b)
    return newlist

##秒数转换
def secchange(secs):
    intlist = []
    newlist = []
    for item in secs:
        intlist.append(int(item))
    for x in intlist:
        if x>=3600:
            hours = int(x/3600)
            minutes = int((x%3600)/60)
            seconds = x%3600%60
            a = '{}小时{}分钟{}秒'.format(hours,minutes,seconds)
            newlist.append(a)
            return newlist
        elif 60<=x<3600:
            minutes = int(x/60)
            seconds = x%60
            a = '{}分钟{}秒'.format(minutes,seconds)
            newlist.append(a)
            return newlist
        elif 0<=x<60:
            a = '{}秒'.format(x)
            newlist.append(a)
            return newlist
        else:
            return 'pasttimewrong'

#list 过滤器
def list_solve1(data_list,x,z,filter_num):
    result_list = []
    filter_list = []
    datasolve_list = []
    a = []
    b = []
    c = []
    d = []
    e = []
    f = []
    g = []
    h = []
    j = []
    k = []
    l = []
    for i in data_list:
        a.append(i[0])
        b.append(i[1])
        c.append(i[2])
        d.append(i[3])
        e.append(i[4])
        f.append(i[5])
        g.append(i[6])
        h.append(i[7])
        j.append(i[8])
        k.append(i[9])
        l.append(i[10])
    datasolve_list.append(a)
    datasolve_list.append(b)
    datasolve_list.append(c)
    datasolve_list.append(d)
    datasolve_list.append(e)
    datasolve_list.append(f)
    datasolve_list.append(g)
    datasolve_list.append(h)
    datasolve_list.append(j)
    datasolve_list.append(k)
    datasolve_list.append(l)
    for num,value in enumerate(datasolve_list[z]):
        if value == filter_num:
            filter_list.append(num)
    if datasolve_list:
        for y in filter_list:
            result_list.append(datasolve_list[x][y])
    return result_list


###队列声明
timequeue = queue.Queue()
myqueue = queue.Queue()
def table(request):
    '''
    st: starttime,
    et: endtime,
    a: membernumber,
    b: paperSpec,
    c: paperWeight,
    d: devicenumber,
    e: pagenumber
    f:  getexcel
    '''
    list_total = gettotalJsonData()
    myqueue.put(list_total)
    if request.method == 'GET':
        if request.GET.get('a'):
            filter_num = str(int(request.GET.get('a')))
            data_list = myqueue.get()
            data_list = list(data_list)
            membernumber = list_solve1(data_list,0,0,filter_num)
            shiftTimestamp = list_solve1(data_list,1,0,filter_num)
            # shiftTimestamp2 = time_change(shiftTimestamp)
            factoryTime = list_solve1(data_list,2,0,filter_num)
            paperAmou = list_solve1(data_list,3,0,filter_num)
            cupAmou = list_solve1(data_list,4,0,filter_num)
            defectiveRate = list_solve1(data_list,5,0,filter_num)
            paperSpec = list_solve1(data_list,6,0,filter_num)
            paperWeight = list_solve1(data_list,7,0,filter_num)
            alarmCount =list_solve1(data_list,8,0,filter_num)
            deviceId = list_solve1(data_list,9,0,filter_num)
            pageId = list_solve1(data_list,10,0,filter_num)
            print(111111111111)
            list_total = list(zip(membernumber,shiftTimestamp,factoryTime,paperAmou,cupAmou,defectiveRate,paperSpec,paperWeight,alarmCount,deviceId,pageId))
            myqueue.queue.clear()
            myqueue.put(list_total)
        if request.GET.get('b'):
            filter_num = str(int(request.GET.get('b')))
            data_list = myqueue.get()
            data_list = list(data_list)
            membernumber = list_solve1(data_list,0,6,filter_num)
            shiftTimestamp = list_solve1(data_list,1,6,filter_num)
            shiftTimestamp2 = time_change(shiftTimestamp)
            factoryTime = list_solve1(data_list,2,6,filter_num)
            paperAmou = list_solve1(data_list,3,6,filter_num)
            cupAmou = list_solve1(data_list,4,6,filter_num)
            defectiveRate = list_solve1(data_list,5,6,filter_num)
            paperSpec = list_solve1(data_list,6,6,filter_num)
            paperWeight = list_solve1(data_list,7,6,filter_num)
            alarmCount =list_solve1(data_list,8,6,filter_num)
            deviceId = list_solve1(data_list,9,6,filter_num)
            pageId = list_solve1(data_list,10,6,filter_num)
            list_total = list(zip(membernumber,shiftTimestamp,factoryTime,paperAmou,cupAmou,defectiveRate,paperSpec,paperWeight,alarmCount,deviceId,pageId))
            myqueue.queue.clear()
            myqueue.put(list_total)
        if request.GET.get('c'):
            filter_num = str(re.findall(r'[-+]?[0-9]*\.?[0-9]+',request.GET.get('c'))[0])
            print(filter_num)
            data_list = myqueue.get()
            data_list = list(data_list)
            membernumber = list_solve1(data_list,0,7,filter_num)
            shiftTimestamp = list_solve1(data_list,1,7,filter_num)
            shiftTimestamp2 = time_change(shiftTimestamp)
            factoryTime = list_solve1(data_list,2,7,filter_num)
            paperAmou = list_solve1(data_list,3,7,filter_num)
            cupAmou = list_solve1(data_list,4,7,filter_num)
            defectiveRate = list_solve1(data_list,5,7,filter_num)
            paperSpec = list_solve1(data_list,6,7,filter_num)
            paperWeight = list_solve1(data_list,7,7,filter_num)
            alarmCount =list_solve1(data_list,8,7,filter_num)
            deviceId = list_solve1(data_list,9,7,filter_num)
            pageId = list_solve1(data_list,10,7,filter_num)
            list_total = list(zip(membernumber,shiftTimestamp,factoryTime,paperAmou,cupAmou,defectiveRate,paperSpec,paperWeight,alarmCount,deviceId,pageId))
            myqueue.queue.clear()
            myqueue.put(list_total)
        if request.GET.get('d'):
            filter_num = str(int(request.GET.get('d')))
            data_list = myqueue.get()
            data_list = list(data_list)
            membernumber = list_solve1(data_list,0,9,filter_num)
            shiftTimestamp = list_solve1(data_list,1,9,filter_num)
            shiftTimestamp2 = time_change(shiftTimestamp)
            factoryTime = list_solve1(data_list,2,9,filter_num)
            paperAmou = list_solve1(data_list,3,9,filter_num)
            cupAmou = list_solve1(data_list,4,9,filter_num)
            defectiveRate = list_solve1(data_list,5,9,filter_num)
            paperSpec = list_solve1(data_list,6,9,filter_num)
            paperWeight = list_solve1(data_list,7,9,filter_num)
            alarmCount =list_solve1(data_list,8,9,filter_num)
            deviceId = list_solve1(data_list,9,9,filter_num)
            pageId = list_solve1(data_list,10,9,filter_num)
            list_total = list(zip(membernumber,shiftTimestamp,factoryTime,paperAmou,cupAmou,defectiveRate,paperSpec,paperWeight,alarmCount,deviceId,pageId))
            myqueue.queue.clear()
            myqueue.put(list_total)
        if request.GET.get('e'):
            filter_num = str(int(request.GET.get('e')))
            data_list = myqueue.get()
            data_list = list(data_list)
            membernumber = list_solve1(data_list,0,10,filter_num)
            shiftTimestamp = list_solve1(data_list,1,10,filter_num)
            shiftTimestamp2 = time_change(shiftTimestamp)
            factoryTime = list_solve1(data_list,2,10,filter_num)
            paperAmou = list_solve1(data_list,3,10,filter_num)
            cupAmou = list_solve1(data_list,4,10,filter_num)
            defectiveRate = list_solve1(data_list,5,10,filter_num)
            paperSpec = list_solve1(data_list,6,10,filter_num)
            paperWeight = list_solve1(data_list,7,10,filter_num)
            alarmCount =list_solve1(data_list,8,10,filter_num)
            deviceId = list_solve1(data_list,9,10,filter_num)
            pageId = list_solve1(data_list,10,10,filter_num)
            list_total = list(zip(membernumber,shiftTimestamp,factoryTime,paperAmou,cupAmou,defectiveRate,paperSpec,paperWeight,alarmCount,deviceId,pageId))
            myqueue.queue.clear()
            myqueue.put(list_total)
        if request.GET.get('st') and request.GET.get('et'):
            default_st = request.GET.get('st')
            d = json.loads(default_st)
            default_et = request.GET.get('et')
            e = json.loads(default_et)
            startTime = datetime.datetime(year = d['year'],month=d['month'],day=d['date'],hour=d['hours'],minute=d['minutes'],second=d['seconds'])
            print(startTime)
            endTime = datetime.datetime(year = e['year'],month=e['month'],day=e['date'],hour=e['hours'],minute=e['minutes'],second=e['seconds'])
            jsonSave(request,startTime,endTime)
            list_total = gettotalJsonData()
            myqueue.queue.clear()
            myqueue.put(list_total)
            timequeue.queue.clear()
            timequeue.put([startTime,endTime],block=False)
        return render(request,'reports/otherbase.html',context={'list_total':list_total})


###excel下载跳转
def test(request):
    timelist = timequeue
    if not timelist.empty():
        timelist = timequeue.get()
        startTime = timelist[0].strftime('%Y年%m月%d日')
        endTime = timelist[-1].strftime('%Y年%m月%d日')
    else:
        startTime = '2017年1月1日'
        endTime = arrow.utcnow().strftime('%Y年%m月%d日')
    filename = ''.join('{0} {1}'.format('状况报告','{0}~{1}'.format(startTime,endTime)))
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        # response =HttpResponse(content_type='application/vnd.ms-excel')
    response['Content-Disposition'] = 'attachment; filename=%s.xlsx' % filename.encode('utf-8').decode('ISO-8859-1')
    row = 0
    col = 0
    colname = ['员工编号', '换班时间', '机器运行时间', '纸片数量', '纸杯数量', '损耗率', '规格', '克重','报警数','设备编号','订单编号']
    workbook = xlsxwriter.Workbook(response, {'in_memory': True})
    format = workbook.add_format()
    worksheet = workbook.add_worksheet()
    format.set_font_color('red')
    # 新建列名
    for item in colname:
        worksheet.write(row, col, item)
        col += 1
    # data = myqueue.get()
    row = 1
    col = 0
    data = myqueue
    if not data.empty():
        datalist = myqueue.get()
        for x in datalist:
            worksheet.write(row,col,x[0])
            worksheet.write(row, col+1, x[1])
            worksheet.write(row, col+2, x[2])
            worksheet.write(row, col+3, x[3])
            worksheet.write(row, col+4, x[4])
            worksheet.write(row, col+5, x[5])
            worksheet.write(row, col+6, x[6])
            worksheet.write(row, col+7, x[7])
            worksheet.write(row, col+8, x[8])
            worksheet.write(row, col+9, x[9])
            worksheet.write(row, col+10, x[10])
            row += 1
            # if float(x[5])>2:
            #     worksheet.write(row,5, x[5])
            # else:
            #     worksheet.write(row, 5, x[5],format)
        # datatuple = tuple(datalist)
        # for a,b,c,d,e,f,g,i,j,k in (datatuple):
        #     worksheet.write(row, col, a)

        #     row += 1
    else:
        return HttpResponse('请输入时间或其他选项')

    workbook.close()
    return response


################
# def vote(request):
#     # return HttpResponse("Hello, world. You're voting from index.")
#     timestart = time.time()
#     report_type = request.POST['type']
#     report_date = request.POST['date']
#     machine = Machine.objects.all().values_list('machine_id')
#     tag = Tag.objects.all().values_list('tag_id')
#     # s = "Report Type: %s; Report Date: %s" % (report_type, report_date)
#     # 从云端获取数据
#     if report_type == 'daily':
#         date_begin_utc = arrow.get(report_date)
#         # 默认时区为0,先减8小时,随后修改时区,匹配本地时间
#         date_begin = date_begin_utc.replace(hours=-8)
#         date_end = date_begin.replace(days=+1)
#         # 请求数据的起始时间
#         startTime = date_begin.to('local')
#         endTime = date_end.to('local')
#         daily = '日报'.encode('gb2312')
#         filename = ''.join('{0} {1}'.format('日报',str(report_date)))
#         # print(filename)
#     elif report_type == 'weekly':
#         report_datetime = datetime.date(int(report_date.split(
#             '-')[0]), int(report_date.split('-')[1]), int(report_date.split('-')[2]))
#         report_datetime_week = int(report_datetime.weekday())
#         date_now = arrow.get(report_date)
#         date_begin_utc = date_now.replace(days=-report_datetime_week)
#         date_begin = date_begin_utc.replace(hours=-8)
#         date_end = date_begin.replace(days=+7)
#         startTime = date_begin.to('local')
#         endTime = date_end.to('local')
#         filename = ''.join('{0} {1}'.format('周报',str(report_date)))
#         # print(filename)
#     elif report_type == 'monthly':
#         report_date_first = report_date[:-2] + '01'
#         date_begin_utc = arrow.get(report_date_first)
#         date_begin = date_begin_utc.replace(hours=-8)
#         date_end = date_begin.replace(months=+1)
#         startTime = date_begin.to('local')
#         endTime = date_end.to('local')
#         print(startTime)
#         print(endTime)
#         filename = ''.join('{0} {1}'.format('月报',str(report_date[:-3])))
#         # print(filename)
#     t0 = arrow.utcnow().float_timestamp
#     # 网关的ID，每个场地都不一样
#     # 获取方法见 08_定制开发看板使用说明_BEDRAGON_V1.0.pdf 3.2.2节
#     # machine_id = "yaYFRz.255.1"
#     machine_ids = []
#     for item in machine:
#         for x in item:
#             machine_ids.append(x)
#     print(machine_ids)
#     tag_ids = []
#     for item in tag:
#         for x in item:
#             tag_ids.append(x)
#     print(tag_ids)
#     # 获取机器种类
#     # 在网关中定义的变量ID，有多少个id就对应多少个料仓
#     # 以50开始，每个料仓递增2
#     # tags = [50, 52, 54, 56, 58, 60, 62, 64, 66, 68, 70, 72]
#     # tags = [21, 22, 23, 24, 25, 26, 27, 28, 29,30]
#
#     # 从云端获取数据
#     r0 = getData(startTime, endTime, machine_ids, tag_ids)
#     # for debug only
#     write_json_file("result_r0.json", r0)
#
#     if r0 is None or 'error' in r0:
#         return HttpResponse("Network Failure!")
#
#     # 转换时间戳格式
#     r1 = convertDateTime(r0)
#     # print(r1)
#
#     # for debug only
#     write_json_file("result_r1.json", r1)
#
#     t1 = arrow.utcnow().float_timestamp
#     delta_t = t1 - t0
#     print("Time cost: %.3f seconds" % delta_t)
#
#     # excel 相关操作
#     response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
#     response['Content-Disposition'] = 'attachment; filename=%s.xlsx' % filename.encode('utf-8').decode('ISO-8859-1')
#     row = 0
#     col = 0
#     colname = ['员工编号', '换班时间', '机器运行时间', '纸片数量', '纸杯数量', '损耗率', '规格', '克重','报警数','设备编号','订单编号']
#     workbook = xlsxwriter.Workbook(response, {'in_memory': True})
#     format = workbook.add_format()
#     worksheet = workbook.add_worksheet()
#     format.set_font_color('red')
#     # 新建列名
#     for item in colname:
#         worksheet.write(row, col, item)
#         col += 1
#     # 数据填充
#     row = 1
#     col = -1
#     row2 = 1
#     col2 = 5
#     del_list = del_repeat_list()
#     for x in [0,1,2,3,4,5,6,7,8,9,10]:
#         datavalue = getJsonData(x)
#         if datavalue:
#             z = 0
#             for y in del_list:
#                 datavalue.pop(y - z)
#                 z += 1
#             col += 1
#             row = 1
#             ####change
#         for item in datavalue:
#             worksheet.write(row, col, item)
#             row += 1
#
#         # for item in datavalue:
#         #     if int(x)<5:
#         #         worksheet.write(row, col, item)
#         #         row += 1
#         #     elif int(x)>5:
#         #         worksheet.write(row2, col2, item)
#         #         row2 += 1
#     #####损耗率超过2红字
#     # row = 1
#     # col = 3
#     # lost = del_repeat_list()
#     # ldatavalue = getJsonData(5)
#     # if ldatavalue:
#     #     z = 0
#     #     for y in del_list:
#     #         ldatavalue.pop(y - z)
#     #         z += 1
#     #     col += 1
#     #     row = 1
#     # for item in ldatavalue:
#     #     if float(item)>2:
#     #         worksheet.write(row, col, item,format)
#     #     else:
#     #         worksheet.write(row, col, item)
#     #     row += 1
#
#
#     # 将shiftTime时间戳列转化为日期格式
#     row = 1
#     col = 2
#     datavalueshifttime = getJsonData(3)
#     z = 0
#     for y in del_list:
#         datavalueshifttime.pop(y - z)
#         z += 1
#     shifttime = []
#     for item in datavalueshifttime:
#         if item == 'None\n':
#             shifttime.append('None')
#         else:
#             localtime = arrow.get(item).to('local')
#             shifttime.append(localtime.format())
#     for item in shifttime:
#         item = item[2:-6]
#         worksheet.write(row, col, item)
#         row += 1
#
#     row = 1
#     col = 1
#     datavalueshifttime = getJsonData(2)
#     z = 0
#     for y in del_list:
#         datavalueshifttime.pop(y - z)
#         z += 1
#     shifttime = []
#     for item in datavalueshifttime:
#         if item == 'None\n':
#             shifttime.append('None')
#         else:
#             localtime = arrow.get(item).to('local')
#             shifttime.append(localtime.format())
#     for item in shifttime:
#         item = item[2:-6]
#         worksheet.write(row, col, item)
#         row += 1
#
#     #求出产量的差值
#     # row = 1
#     # col = 8
#     # datavalueproductcount = getJsonData(9)
#     # z = 0
#     # if datavalueproductcount:
#     #     for y in del_list:
#     #         datavalueproductcount.pop(y - z)
#     #         z += 1
#     #     productcount = []
#     #     for num,item in enumerate(datavalueproductcount):
#     #         if item == 'None\n' or item == '0\n':
#     #             datavalueproductcount[num] = '0\n'
#     #             productcount.append('0\n')
#     #         else:
#     #             if num>0:
#     #                 if int(datavalueproductcount[num-1][:-1])!= 0:
#     #                     productcount.append(int(datavalueproductcount[num][:-1])-(int(datavalueproductcount[num-1][:-1])))
#     #                 else:
#     #                     productcount.append(int(datavalueproductcount[num][:-1])-int(datavalueproductcount[num-2][:-1]))
#     #     for item in productcount:
#     #         if item == '0\n':
#     #             item = 'None'
#     #         worksheet.write(row, col, item)
#     #         row += 1
#     #     # print(productcount)
#     #     productcount.append('...')
#     workbook.close()
#     timeover = time.time()
#     print(timeover-timestart)
#     return response
#
# def del_repeat():
#     #删除flag重复项以及所对应的其他参数
#     pass


