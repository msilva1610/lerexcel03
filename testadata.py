import datetime 
from dateutil.parser import parse
from datetime import date

def main(): 
    # dt_stamp = datetime.datetime(2019, 1, 9)
    # dt_stamp = "2018-05-28 00:00:00"

    # print(type(dt_stamp))

    # print(str(dt_stamp))

    # print(dt_stamp.strftime('%Y-%m-%d'))

    str_stamp = '2019-01-09 00:00:00'
    str_stamp = '2019-01-10 '
    str_stamp = None
    str_stamp = '01/01/2020 02/01/2020 '

    # print(datetime.datetime.strptime(str_stamp, '%Y-%m-%d %H:%M:%S'))
    # minhadata = (datetime.datetime.strptime(str_stamp, '%Y-%m-%d %H:%M:%S'))
    # print(type(minhadata))

    # print(minhadata.strftime('%Y-%m-%d'))

    #isinstance(now, datetime.datetime)

    # print(convertData(str_stamp))
    # print(convertDataParse(str_stamp))
    
    # organizaDatas()
    # classificadatas()    
    
    iniciodev = "24/08/2018 28/09/2018 30/11/2018 06/12/2018 07/12/2018 16/01/2019 "

    print(tratadatas(iniciodev))

def convertData(strdata):
    data = strdata
    try:
        minhadata = (datetime.datetime.strptime(data, '%Y-%m-%d %H:%M:%S'))
    except:
        minhadata = strdata
    return data

def convertDataParse(strdata):
    data = strdata
    try:
        if not isinstance(data, datetime.datetime):
            dataconvertida = parse(data)
            data = dataconvertida.strftime('%d/%m/%Y')
        else:
            data = dataconvertida.strftime('%d/%m/%Y')
    except:
        data = strdata
    return data


def organizaDatas():
    iniciodev = "24/08/2018 28/09/2018 30/11/2018 06/12/2018 07/12/2018 16/01/2019 "
    arrdatas = iniciodev.split(' ')
    arrdatas.sort()
    print(arrdatas)

def classificadatas():
    retorno = {}
    iniciodev = "24/08/2018 28/09/2018 30/11/2018 06/12/2018 07/12/2018 16/01/2019 "
    datas = iniciodev.split(' ')
    arr = []
    dt = None
    for d in datas:
        if len(d) >= 10: 
            dt = datetime.datetime.strptime(d, '%d/%m/%Y')
            arr.append(dt)
            arr.sort()
    for d in arr:
        dataconvertida = d.strftime('%d/%m/%Y')
        # print(dataconvertida)
    maiorData = len(arr) - 1   

    d1 = arr[0].strftime('%d/%m/%Y')
    d2 = arr[maiorData].strftime('%d/%m/%Y')
    diff  = parse(d2) - parse(d1)

    retorno.update({'Menor data': arr[0].strftime('%d/%m/%Y')})
    retorno.update({'Maior data': arr[maiorData].strftime('%d/%m/%Y')})
    retorno.update({'Delta': diff.days})

    print(retorno)

def tratadatas(arrDatas):
    menorData = None
    maiorData = None
    delta = 0
    lstdatas = []
    try:
        datas = arrDatas.split(' ')
        for d in datas:
            try:
                data = datetime.datetime.strptime(d, '%d/%m/%Y')
                lstdatas.append(data)
                lstdatas.sort()
            except:
                continue
        print(lstdatas)
        menorData = lstdatas[0]    
        maiorData = len(lstdatas) - 1   
        d1 = lstdatas[0].strftime('%d/%m/%Y')
        d2 = lstdatas[maiorData].strftime('%d/%m/%Y')
        delta  = parse(d2) - parse(d1)
        return lstdatas[0].strftime('%d/%m/%Y'),lstdatas[maiorData].strftime('%d/%m/%Y'),delta.days
    except:
        pass
    # return lstdatas[0].strftime('%d/%m/%Y'),lstdatas[maiorData].strftime('%d/%m/%Y'),delta.days

main()