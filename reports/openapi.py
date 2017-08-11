import json
import hashlib
import arrow
import requests
# import random
from .models import Config, Machine, Tag


def getData():
    config = Config.objects.order_by('id')[:1][0]
    customerCode = config.customer_code
    appkey = config.app_key
    appsecret = config.app_secret
    apiurl = config.api_url
    headers = {}
    headers['Content-Type'] = 'application/json'
    headers['X-Auth-AppKey'] = appkey


def getRawData(startTime, endTime, machines, tags, authObject):
    # get the raw data from OPENAPI
    # input paramters:
    # startTime = '2017-03-08T08:00:00+08:00'
    # endTime = '2017-03-08T10:00:00+08:00'
    # machines = ["yaYFRz.255.1"]
    # tags = [50, 52, 54, 56, 58, 60, 62, 64, 66, 68, 70, 72]
    # authObject['customerCode'] = "chinasks"
    # authObject['appkey'] = "ZMPZ6hl60UW43exaSvoPTg=="
    # authObject['appsecret'] = "84Y17hCX1UyvRysIYML99w=="
    # authObject['apiurl'] = 'http://openapi.energymost.com/API/Energy.svc/GetEnergyUsageData'
    # -------------------------------------------------------------------------
    headers = {}
    headers['Content-Type'] = 'application/json'
    headers['X-Auth-AppKey'] = authObject['appkey']

    payload = {}
    payload['CustomerCode'] = authObject['customerCode']
    payload['ProcessPrecision'] = True
    payload['TagCodes'] = []
    for machine_id in machines:
        for tag_id in tags:
            read_tag = str(machine_id) + "." + str(tag_id)
            payload['TagCodes'].append(read_tag)
    print(payload['TagCodes'])
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
    s = authObject['appkey'] + body + authObject['appsecret']
    m.update(s.encode('utf-8'))
    appfig = m.hexdigest()
    # print(appkey)
    # print(body)
    # print(appsecret)
    # print(appfig)
    headers['X-Auth-Fig'] = appfig

    try:
        r = requests.post(authObject['apiurl'], data=body, headers=headers, timeout=10)
    except Exception as e:
        print(e)
        return None

    if r.status_code != 200:
        info = "get data from openAPI failed"
        print(info)
        return None

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
