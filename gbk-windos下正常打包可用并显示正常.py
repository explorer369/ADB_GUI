#_*_coding:GBK
import requests, xlrd, MySQLdb, time, sys,os
from xlutils import copy #从xlutils模块中导入copy这个函数
import encodings.gb18030
localpath = os.getcwd()
sys.path.append(localpath)

# 此脚本已调好，写BUG的还未调试，开关在84行

def readExcel(file_path):
    '''
    读取excel测试用例的函数
    :param file_path:传入一个excel文件，或者文件的绝对路径
    :return:返回这个excel第一个sheet页中的所有测试用例的list
    '''
    try:
        book = xlrd.open_workbook(file_path)               #打开excel
    except Exception,e:                                    #如果路径不在或者excel不正确，返回报错信息  
        print '路径不在或者excel不正确',e
        return e
    else:
        sheet = book.sheet_by_index(0)                     #取第一个sheet页
        rows= sheet.nrows                                  #取这个sheet页的所有行数
        case_list = []                                     #保存每一条case
        for i in range(rows):
            if i !=0:                                      #把每一条测试用例添加到case_list中              
                case_list.append(sheet.row_values(i))
                                                           #调用接口测试的函数，把大的放所有case的list和case的路径传进去，因为后面还需要把返回报文和测试结果写到excel中
        interfaceTest(case_list,file_path)

def interfaceTest(case_list,file_path):
    res_flags = []                       #存测试结果的list    
    request_urls = []                    #存请求报文的list    
    responses = []                       #存返回报文的list
    oldparam = []                        #存储入参值，供后面切割用（）
    
    for case in case_list:
        '''
        先遍历excel中每一条case的值，然后根据对应的索引取到case中每个字段的值
        '''
        try:
            '''
            这里捕捉一下异常，如果excel格式不正确的话，就返回异常
            '''
            
            product = case[0]            # 项目，提bug的时候可以根据项目来提        
            case_id = case[1]            # 用例id，提bug的时候用            
            interface_name = case[2]     # 接口名称，也是提bug的时候用           
            case_detail = case[3]        # 用例描述          
            method = case[4]             # 请求方式           
            url = case[5]                # 请求url          
            param = case[6]              # 入参      
            res_check = case[7]          # 预期结果           
            tester = case[10]            # 测试人员
        
        except Exception,e:
            return '测试用例格式不正确！%s'%e
        if param== '':                           # 如果请求参数是空的话，请求报文就是url，然后把请求报文存到请求报文list中
            new_url = url                        # 请求报文
            request_urls.append(new_url)
        else:          
            # 如果请求参数不为空的话，请求报文就是url+?+参数，格式和下面一样http://127.0.0.1:8080/rest/login?oper_no=marry&id=100，然后把请求报文存到请求报文list中
            
            new_url = url+'?'+urlParam(param)     #请求报文
            
            '''
            excel里面的如果有多个入参的话，参数是用;隔开，a=1;b=2这样的，请求的时候多个参数要用&连接，
            要把;替换成&，所以调用了urlParam这个函数，把参数中的;替换成&，函数在下面定义的
            '''
            
            request_urls.append(new_url)
            oldparam.append(param)           # 把param值保存在oldparam里
            
        if method.upper() == 'GET':    # 如果是get请求就调用requests模块的get方法，.text是获取返回报文，保存返回报文，把返回报文存到返回报文的list中
            print new_url              # 输出请求报文
            results = requests.get(new_url).text
            print results              # 输出返回结果信息           
            responses.append(results)          # 把返回的结果信息添加到responses里，上面定义了responses为空
            res = readRes(results,res_check)   # 获取到返回报文之后需要根据预期结果去判断测试是否通过，调用查看结果方法把返回报文和预期结果传进去，判断是否通过，readRes方法在下面定义了。
        else:                          # 如果不是get请求，也就是post请求，就调用requests模块的post方法，.text是获取返回报文，保存返回报文，把返回报文存到返回报文的list中

            results = requests.post(new_url).text
            responses.append(results)  # 获取到返回报文之后需要根据预期结果去判断测试是否通过，调用查看结果方法把返回报文和预期结果传进去，判断是否通过，readRes方法会返回测试结果，如果返回pass就说明测试通过了，readRes方法在下面定义了。
            res = readRes(results,res_check)  #readRes调用下面的函数（readRes函数中对比后返回结果，对比OK就return了'pass'的）
        if 'pass' in res:              # 判断测试结果，然后把通过或者失败插入到测试结果的list中
            res_flags.append('pass')   # 写入到新的测试结果的EXCL表中，这里是固定写死了的，具体看150行左右，也就是测试结果列写列了
        else:
            res_flags.append('fail')   # 如果不通过，也把结果FAIL写入到EXCL中的测试结果中 
            #writeBug(case_id,interface_name,new_url,results,res_check) # 如果不通过的话，就调用写bug的方法，把case_id、接口名称、请求报文、返回报文和预期结果传进去writeBug方法在下面定义了，具体实现是先连接数据库，然后拼sql，插入到bug表中
    '''
    全部用例执行完之后，会调用copy_excel方法，把测试结果写到excel中，
    每一条用例的请求报文、返回报文、测试结果，这三个每个我在上面都定义了一个list
    来存每一条用例执行的结果，把源excel用例的路径和三个list传进去调用即可，copy_excel方
    法在下面定义了，也加了注释
    '''
    copy_excel(file_path,res_flags,request_urls,responses,oldparam)  # 我在这里多加了一个oldparam，这样下面的函数才可以调用此值

def readRes(res,res_check):
    '''
    :param res: 返回报文
    :param res_check: 预期结果
    :return: 通过或者不通过，不通过的话会把哪个参数和预期不一致返回
    '''
    '''
    返回报文的例子是这样的{"id":"J_775682","p":275.00,"m":"458.00"}
    excel预期结果中的格式是xx=11;xx=22这样的，所以要把返回报文改成xx=22这样的格式
    所以用到字符串替换，把返回报文中的":"和":替换成=，返回报文就变成
    {"id=J_775682","p=275.00,"m=458.00"},这样就和预期结果一样了,当然也可以用python自带的
    json模块来解析json串，但是有的返回的不是标准的json格式，处理起来比较麻烦，这里我就用字符串的方法了
    '''
    res = res.replace('":"',"=").replace('":',"=")   # 把excel预期结果冒号替换为=号

    '''
    res_check是excel中的预期结果，是xx=11;xx=22这样的
    所以用split分割字符串，split是python内置函数，切割字符串，变成一个list
    ['xx=1','xx=2']这样的，然后遍历这个list，判断list中的每个元素是否存在这个list中，
    如果每个元素都在返回报文中的话，就说明和预期结果一致
    上面我们已经把返回报文变成{"id=J_775682","p=275.00,"m=458.00"}
    '''
    res_check = res_check.split(';')    # 切割后的excel预期结果[u'id=3948354', u'title=Python\u5b66\u4e60\u624b\u518c']

    for s in res_check:                 # 遍历这个excel的每一条预期结果
        '''
        遍历预期结果的list，如果在返回报文中，什么都不做，pass代表什么也不做，全部都存在的话，就返回pass
        如果不在的话，就返回错误信息和不一致的字段，因为res_check是从excel里面读出来的
        字符Unicode类型的的，python的字符串是str类型的，所以要用str方法强制类型转换，转换成string类型的
        '''
        if s in res:   # 如果每条excel预期结果在RES里
            pass
        else:
            return   '错误，返回参数和预期结果不一致'+str(s)
    return 'pass'
    
# 参数转换，把参数转换为'xx=11&xx=2这样'    
def urlParam(param):     
    return param.replace(';','&')
def copy_excel(file_path,res_flags,request_urls,responses,oldparam):   # 多了一个oldparam，调用上面赋的值
    '''
    :param file_path: 测试用例的路径
    :param res_flags: 测试结果的list
    :param request_urls: 请求报文的list
    :param responses: 返回报文的list
    :return:
    '''
    '''
    这个函数的作用是写excel，把请求报文、返回报文和测试结果写到测试用例的excel中
    因为xlrd模块只能读excel，不能写，所以用xlutils这个模块，但是python中没有一个模块能
    直接操作已经写好的excel，所以只能用xlutils模块中的copy方法，copy一个新的excel，才能操作
    '''
    
#下面两段是我调试时用的，后来没有用上，保留以后查看方式   
    #oldparam = str(oldparam)                         # 把oldparam转化为字符串（原来是列表）
    #Keyword = oldparam.split("=")[1].split(";")[0]   # 提取=号后面的关键字来作为后面返回报文的截取关键字
    excelpath = localpath + '\\' + 'test_case.xls'
    book = xlrd.open_workbook(excelpath)     #打开原来的excel，获取到这个book对象    
    new_book = copy.copy(book)               #复制一个new_book    
    sheet = new_book.get_sheet(0)            #然后获取到这个复制的excel的第一个sheet页
    i = 1
    for request_url,response,flag in zip(request_urls,responses,res_flags):
        '''
        同时遍历请求报文、返回报文和测试结果这3个大的list
        然后把每一条case执行结果写到excel中，zip函数可以将多个list放在一起遍历
        因为第一行是表头，所以从第二行开始写，也就是索引位1的位置，i代表行
        所以i赋值为1，然后每写一条，然后i+1， i+=1同等于i=i+1
        请求报文、返回报文、测试结果分别在excel的8、9、11列，列是固定的，所以就给写死了
        后面跟上要写的值，因为excel用的是Unicode字符编码，所以前面带个u表示用Unicode编码
        否则会有乱码
        '''
        sheet.write(i,8,u'%s'%request_url)     # 在第8列写入请求报文的list
        if len(response)<100:                  # 如果返回报文小于100，就把返回报文写入到EXCL结果中
            sheet.write(i,9,u'%s'%response)    # 在第9列写入返回报文的list  如果返回的报文太长会报错的，所有有时候就把这条给注释了，但无法看到返回的报文
        else:
            sheet.write(i,9,u'返回的报文太长了，保存时会出错，故未保存')
        sheet.write(i,11,u'%s'%flag)           # 在第11列写入测试结果的list
        i+=1
    #写完之后在当前目录下(可以自己指定一个目录)保存一个以当前时间命名的测试结果，time.strftime()是格式化日期
    new_book.save(u'%s_测试结果.xls'%time.strftime('%Y%m%d%H%M%S'))

def writeBug(bug_id,interface_name,request,response,res_check):
    '''
    这个函数用来连接数据库，往bugfree数据中插入bug，拼sql，执行sql即可
    :param bug_id: bug序号
    :param interface_name: 接口名称
    :param request: 请求报文
    :param response: 返回报文
    :param res_check: 预期结果
    :return:
    '''
    bug_id = bug_id.encode('utf-8')
    interface_name = interface_name.encode('utf-8')
    res_check = res_check.encode('utf-8')
    response = response.encode('utf-8')
    request = request.encode('utf-8')
    '''
    因为上面几个字符串是从excel里面读出来的都是Unicode字符集编码的，
    python的字符串上面指定了utf-8编码的，所以要把它的字符集改成utf-8，才能把sql拼起来
    encode方法可以指定字符集
    '''
    
    now = time.strftime("%Y-%m-%d %H:%M:%S")      # 取当前时间，作为提bug的时间
    
    bug_title = bug_id + '_' + interface_name + '_结果和预期不符'     # bug标题用bug编号加上接口名称然后加上_结果和预期不符，可以自己随便定义要什么样的bug标题
    
    step = '[请求报文]<br />'+request+'<br/>'+'[预期结果]<br/>'+res_check+'<br/>'+'<br/>'+'[响应报文]<br />'+'<br/>'+response     # 复现步骤就是请求报文+预期结果+返回报文
    #拼sql，这里面的项目id，创建人，严重程度，指派给谁，都在sql里面写死，使用的时候可以根据项目和接口
    # 来判断提bug的严重程度和提交给谁
    sql = "INSERT INTO `bf_bug_info` (`created_at`, `created_by`, `updated_at`, `updated_by`, `bug_status`, `assign_to`, `title`, `mail_to`, `repeat_step`, `lock_version`, `resolved_at`, `resolved_by`, `closed_at`, `closed_by`, `related_bug`, `related_case`, `related_result`, " \
          "`productmodule_id`, `modified_by`, `solution`, `duplicate_id`, `product_id`, " \
          "`reopen_count`, `priority`, `severity`) VALUES ('%s', '1', '%s', '1', 'Active', '1', '%s', '系统管理员', '%s', '1', NULL , NULL, NULL, NULL, '', '', '', NULL, " \
          "'1', NULL, NULL, '1', '0', '1', '1');"%(now,now,bug_title,step)
    #建立连接，使用MMySQLdb模块的connect方法连接mysql，传入账号、密码、数据库、端口、ip和字符集
    coon = MySQLdb.connect(user='zzf',passwd='123456',db='bugfree',port=3306,host='127.0.0.1',charset='utf8')
    #建立游标
    cursor = coon.cursor()
    #执行sql
    cursor.execute(sql)
    #提交
    coon.commit()
    #关闭游标
    cursor.close()
    #关闭连接
    coon.close()


if __name__ == '__main__':
    try:
        filename = 'test_case.xls'
        print filename
        readExcel(filename)
        print '执行完毕，请查看测试结果！'
        thisIsLove = input('按 ENTER 键退出窗口！')
    except Exception,e:
        print e