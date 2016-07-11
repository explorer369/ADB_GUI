#_*_coding:GBK
import ntplib,os,datetime,time
def EffectiveTime():
    '''联网验证过期日期
    '''
    # 设置过期日期（注意格式要一致）
    ExpirationTime = '2016-08-09' 
    ExpirationTime = time.strptime(ExpirationTime,'%Y-%m-%d') #转为<type 'time.struct_time'>类型的格式
    #获取网络时间                 
    c = ntplib.NTPClient()
    try:
        response = c.request('pool.ntp.org')
        ts = response.tx_time
        _date = time.strftime('%Y-%m-%d',time.localtime(ts)) #结果：2016-07-10
        _time = time.strftime('%X',time.localtime(ts))      #结果：13:14:26    
        NetworkTime = time.strptime(_date,'%Y-%m-%d')  #获取的时间为str转为<type 'time.struct_time'>
    except Exception,e:
        print e,'\n请联接网络后重试' #\n换行再显示后面的内容
    #获取本地系统时间
    SystemTime = time.strftime("%Y-%m-%d",time.localtime(time.time()))
    SystemTime = time.strptime(SystemTime,'%Y-%m-%d')
    try:
        if ExpirationTime > SystemTime and ExpirationTime > NetworkTime:
            return True
        elif ExpirationTime > NetworkTime:
            return True
        else:
            print '认证过期，请联系作者！'
            return False            
    except (UnboundLocalError,Exception) as e:
        print e,'认证过期，请联系作者！'
        return False