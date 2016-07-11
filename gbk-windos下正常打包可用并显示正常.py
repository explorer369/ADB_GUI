#_*_coding:GBK
import requests, xlrd, MySQLdb, time, sys,os
from xlutils import copy #��xlutilsģ���е���copy�������
import encodings.gb18030
localpath = os.getcwd()
sys.path.append(localpath)

# �˽ű��ѵ��ã�дBUG�Ļ�δ���ԣ�������84��

def readExcel(file_path):
    '''
    ��ȡexcel���������ĺ���
    :param file_path:����һ��excel�ļ��������ļ��ľ���·��
    :return:�������excel��һ��sheetҳ�е����в���������list
    '''
    try:
        book = xlrd.open_workbook(file_path)               #��excel
    except Exception,e:                                    #���·�����ڻ���excel����ȷ�����ر�����Ϣ  
        print '·�����ڻ���excel����ȷ',e
        return e
    else:
        sheet = book.sheet_by_index(0)                     #ȡ��һ��sheetҳ
        rows= sheet.nrows                                  #ȡ���sheetҳ����������
        case_list = []                                     #����ÿһ��case
        for i in range(rows):
            if i !=0:                                      #��ÿһ������������ӵ�case_list��              
                case_list.append(sheet.row_values(i))
                                                           #���ýӿڲ��Եĺ������Ѵ�ķ�����case��list��case��·������ȥ����Ϊ���滹��Ҫ�ѷ��ر��ĺͲ��Խ��д��excel��
        interfaceTest(case_list,file_path)

def interfaceTest(case_list,file_path):
    res_flags = []                       #����Խ����list    
    request_urls = []                    #�������ĵ�list    
    responses = []                       #�淵�ر��ĵ�list
    oldparam = []                        #�洢���ֵ���������и��ã���
    
    for case in case_list:
        '''
        �ȱ���excel��ÿһ��case��ֵ��Ȼ����ݶ�Ӧ������ȡ��case��ÿ���ֶε�ֵ
        '''
        try:
            '''
            ���ﲶ׽һ���쳣�����excel��ʽ����ȷ�Ļ����ͷ����쳣
            '''
            
            product = case[0]            # ��Ŀ����bug��ʱ����Ը�����Ŀ����        
            case_id = case[1]            # ����id����bug��ʱ����            
            interface_name = case[2]     # �ӿ����ƣ�Ҳ����bug��ʱ����           
            case_detail = case[3]        # ��������          
            method = case[4]             # ����ʽ           
            url = case[5]                # ����url          
            param = case[6]              # ���      
            res_check = case[7]          # Ԥ�ڽ��           
            tester = case[10]            # ������Ա
        
        except Exception,e:
            return '����������ʽ����ȷ��%s'%e
        if param== '':                           # �����������ǿյĻ��������ľ���url��Ȼ��������Ĵ浽������list��
            new_url = url                        # ������
            request_urls.append(new_url)
        else:          
            # ������������Ϊ�յĻ��������ľ���url+?+��������ʽ������һ��http://127.0.0.1:8080/rest/login?oper_no=marry&id=100��Ȼ��������Ĵ浽������list��
            
            new_url = url+'?'+urlParam(param)     #������
            
            '''
            excel���������ж����εĻ�����������;������a=1;b=2�����ģ������ʱ��������Ҫ��&���ӣ�
            Ҫ��;�滻��&�����Ե�����urlParam����������Ѳ����е�;�滻��&�����������涨���
            '''
            
            request_urls.append(new_url)
            oldparam.append(param)           # ��paramֵ������oldparam��
            
        if method.upper() == 'GET':    # �����get����͵���requestsģ���get������.text�ǻ�ȡ���ر��ģ����淵�ر��ģ��ѷ��ر��Ĵ浽���ر��ĵ�list��
            print new_url              # ���������
            results = requests.get(new_url).text
            print results              # ������ؽ����Ϣ           
            responses.append(results)          # �ѷ��صĽ����Ϣ��ӵ�responses����涨����responsesΪ��
            res = readRes(results,res_check)   # ��ȡ�����ر���֮����Ҫ����Ԥ�ڽ��ȥ�жϲ����Ƿ�ͨ�������ò鿴��������ѷ��ر��ĺ�Ԥ�ڽ������ȥ���ж��Ƿ�ͨ����readRes���������涨���ˡ�
        else:                          # �������get����Ҳ����post���󣬾͵���requestsģ���post������.text�ǻ�ȡ���ر��ģ����淵�ر��ģ��ѷ��ر��Ĵ浽���ر��ĵ�list��

            results = requests.post(new_url).text
            responses.append(results)  # ��ȡ�����ر���֮����Ҫ����Ԥ�ڽ��ȥ�жϲ����Ƿ�ͨ�������ò鿴��������ѷ��ر��ĺ�Ԥ�ڽ������ȥ���ж��Ƿ�ͨ����readRes�����᷵�ز��Խ�����������pass��˵������ͨ���ˣ�readRes���������涨���ˡ�
            res = readRes(results,res_check)  #readRes��������ĺ�����readRes�����жԱȺ󷵻ؽ�����Ա�OK��return��'pass'�ģ�
        if 'pass' in res:              # �жϲ��Խ����Ȼ���ͨ������ʧ�ܲ��뵽���Խ����list��
            res_flags.append('pass')   # д�뵽�µĲ��Խ����EXCL���У������ǹ̶�д���˵ģ����忴150�����ң�Ҳ���ǲ��Խ����д����
        else:
            res_flags.append('fail')   # �����ͨ����Ҳ�ѽ��FAILд�뵽EXCL�еĲ��Խ���� 
            #writeBug(case_id,interface_name,new_url,results,res_check) # �����ͨ���Ļ����͵���дbug�ķ�������case_id���ӿ����ơ������ġ����ر��ĺ�Ԥ�ڽ������ȥwriteBug���������涨���ˣ�����ʵ�������������ݿ⣬Ȼ��ƴsql�����뵽bug����
    '''
    ȫ������ִ����֮�󣬻����copy_excel�������Ѳ��Խ��д��excel�У�
    ÿһ�������������ġ����ر��ġ����Խ����������ÿ���������涼������һ��list
    ����ÿһ������ִ�еĽ������Դexcel������·��������list����ȥ���ü��ɣ�copy_excel��
    �������涨���ˣ�Ҳ����ע��
    '''
    copy_excel(file_path,res_flags,request_urls,responses,oldparam)  # ������������һ��oldparam����������ĺ����ſ��Ե��ô�ֵ

def readRes(res,res_check):
    '''
    :param res: ���ر���
    :param res_check: Ԥ�ڽ��
    :return: ͨ�����߲�ͨ������ͨ���Ļ�����ĸ�������Ԥ�ڲ�һ�·���
    '''
    '''
    ���ر��ĵ�������������{"id":"J_775682","p":275.00,"m":"458.00"}
    excelԤ�ڽ���еĸ�ʽ��xx=11;xx=22�����ģ�����Ҫ�ѷ��ر��ĸĳ�xx=22�����ĸ�ʽ
    �����õ��ַ����滻���ѷ��ر����е�":"��":�滻��=�����ر��ľͱ��
    {"id=J_775682","p=275.00,"m=458.00"},�����ͺ�Ԥ�ڽ��һ����,��ȻҲ������python�Դ���
    jsonģ��������json���������еķ��صĲ��Ǳ�׼��json��ʽ�����������Ƚ��鷳�������Ҿ����ַ����ķ�����
    '''
    res = res.replace('":"',"=").replace('":',"=")   # ��excelԤ�ڽ��ð���滻Ϊ=��

    '''
    res_check��excel�е�Ԥ�ڽ������xx=11;xx=22������
    ������split�ָ��ַ�����split��python���ú������и��ַ��������һ��list
    ['xx=1','xx=2']�����ģ�Ȼ��������list���ж�list�е�ÿ��Ԫ���Ƿ�������list�У�
    ���ÿ��Ԫ�ض��ڷ��ر����еĻ�����˵����Ԥ�ڽ��һ��
    ���������Ѿ��ѷ��ر��ı��{"id=J_775682","p=275.00,"m=458.00"}
    '''
    res_check = res_check.split(';')    # �и���excelԤ�ڽ��[u'id=3948354', u'title=Python\u5b66\u4e60\u624b\u518c']

    for s in res_check:                 # �������excel��ÿһ��Ԥ�ڽ��
        '''
        ����Ԥ�ڽ����list������ڷ��ر����У�ʲô��������pass����ʲôҲ������ȫ�������ڵĻ����ͷ���pass
        ������ڵĻ����ͷ��ش�����Ϣ�Ͳ�һ�µ��ֶΣ���Ϊres_check�Ǵ�excel�����������
        �ַ�Unicode���͵ĵģ�python���ַ�����str���͵ģ�����Ҫ��str����ǿ������ת����ת����string���͵�
        '''
        if s in res:   # ���ÿ��excelԤ�ڽ����RES��
            pass
        else:
            return   '���󣬷��ز�����Ԥ�ڽ����һ��'+str(s)
    return 'pass'
    
# ����ת�����Ѳ���ת��Ϊ'xx=11&xx=2����'    
def urlParam(param):     
    return param.replace(';','&')
def copy_excel(file_path,res_flags,request_urls,responses,oldparam):   # ����һ��oldparam���������渳��ֵ
    '''
    :param file_path: ����������·��
    :param res_flags: ���Խ����list
    :param request_urls: �����ĵ�list
    :param responses: ���ر��ĵ�list
    :return:
    '''
    '''
    ���������������дexcel���������ġ����ر��ĺͲ��Խ��д������������excel��
    ��Ϊxlrdģ��ֻ�ܶ�excel������д��������xlutils���ģ�飬����python��û��һ��ģ����
    ֱ�Ӳ����Ѿ�д�õ�excel������ֻ����xlutilsģ���е�copy������copyһ���µ�excel�����ܲ���
    '''
    
#�����������ҵ���ʱ�õģ�����û�����ϣ������Ժ�鿴��ʽ   
    #oldparam = str(oldparam)                         # ��oldparamת��Ϊ�ַ�����ԭ�����б�
    #Keyword = oldparam.split("=")[1].split(";")[0]   # ��ȡ=�ź���Ĺؼ�������Ϊ���淵�ر��ĵĽ�ȡ�ؼ���
    excelpath = localpath + '\\' + 'test_case.xls'
    book = xlrd.open_workbook(excelpath)     #��ԭ����excel����ȡ�����book����    
    new_book = copy.copy(book)               #����һ��new_book    
    sheet = new_book.get_sheet(0)            #Ȼ���ȡ��������Ƶ�excel�ĵ�һ��sheetҳ
    i = 1
    for request_url,response,flag in zip(request_urls,responses,res_flags):
        '''
        ͬʱ���������ġ����ر��ĺͲ��Խ����3�����list
        Ȼ���ÿһ��caseִ�н��д��excel�У�zip�������Խ����list����һ�����
        ��Ϊ��һ���Ǳ�ͷ�����Դӵڶ��п�ʼд��Ҳ��������λ1��λ�ã�i������
        ����i��ֵΪ1��Ȼ��ÿдһ����Ȼ��i+1�� i+=1ͬ����i=i+1
        �����ġ����ر��ġ����Խ���ֱ���excel��8��9��11�У����ǹ̶��ģ����Ծ͸�д����
        �������Ҫд��ֵ����Ϊexcel�õ���Unicode�ַ����룬����ǰ�����u��ʾ��Unicode����
        �����������
        '''
        sheet.write(i,8,u'%s'%request_url)     # �ڵ�8��д�������ĵ�list
        if len(response)<100:                  # ������ر���С��100���Ͱѷ��ر���д�뵽EXCL�����
            sheet.write(i,9,u'%s'%response)    # �ڵ�9��д�뷵�ر��ĵ�list  ������صı���̫���ᱨ��ģ�������ʱ��Ͱ�������ע���ˣ����޷��������صı���
        else:
            sheet.write(i,9,u'���صı���̫���ˣ�����ʱ�������δ����')
        sheet.write(i,11,u'%s'%flag)           # �ڵ�11��д����Խ����list
        i+=1
    #д��֮���ڵ�ǰĿ¼��(�����Լ�ָ��һ��Ŀ¼)����һ���Ե�ǰʱ�������Ĳ��Խ����time.strftime()�Ǹ�ʽ������
    new_book.save(u'%s_���Խ��.xls'%time.strftime('%Y%m%d%H%M%S'))

def writeBug(bug_id,interface_name,request,response,res_check):
    '''
    ������������������ݿ⣬��bugfree�����в���bug��ƴsql��ִ��sql����
    :param bug_id: bug���
    :param interface_name: �ӿ�����
    :param request: ������
    :param response: ���ر���
    :param res_check: Ԥ�ڽ��
    :return:
    '''
    bug_id = bug_id.encode('utf-8')
    interface_name = interface_name.encode('utf-8')
    res_check = res_check.encode('utf-8')
    response = response.encode('utf-8')
    request = request.encode('utf-8')
    '''
    ��Ϊ���漸���ַ����Ǵ�excel����������Ķ���Unicode�ַ�������ģ�
    python���ַ�������ָ����utf-8����ģ�����Ҫ�������ַ����ĳ�utf-8�����ܰ�sqlƴ����
    encode��������ָ���ַ���
    '''
    
    now = time.strftime("%Y-%m-%d %H:%M:%S")      # ȡ��ǰʱ�䣬��Ϊ��bug��ʱ��
    
    bug_title = bug_id + '_' + interface_name + '_�����Ԥ�ڲ���'     # bug������bug��ż��Ͻӿ�����Ȼ�����_�����Ԥ�ڲ����������Լ���㶨��Ҫʲô����bug����
    
    step = '[������]<br />'+request+'<br/>'+'[Ԥ�ڽ��]<br/>'+res_check+'<br/>'+'<br/>'+'[��Ӧ����]<br />'+'<br/>'+response     # ���ֲ������������+Ԥ�ڽ��+���ر���
    #ƴsql�����������Ŀid�������ˣ����س̶ȣ�ָ�ɸ�˭������sql����д����ʹ�õ�ʱ����Ը�����Ŀ�ͽӿ�
    # ���ж���bug�����س̶Ⱥ��ύ��˭
    sql = "INSERT INTO `bf_bug_info` (`created_at`, `created_by`, `updated_at`, `updated_by`, `bug_status`, `assign_to`, `title`, `mail_to`, `repeat_step`, `lock_version`, `resolved_at`, `resolved_by`, `closed_at`, `closed_by`, `related_bug`, `related_case`, `related_result`, " \
          "`productmodule_id`, `modified_by`, `solution`, `duplicate_id`, `product_id`, " \
          "`reopen_count`, `priority`, `severity`) VALUES ('%s', '1', '%s', '1', 'Active', '1', '%s', 'ϵͳ����Ա', '%s', '1', NULL , NULL, NULL, NULL, '', '', '', NULL, " \
          "'1', NULL, NULL, '1', '0', '1', '1');"%(now,now,bug_title,step)
    #�������ӣ�ʹ��MMySQLdbģ���connect��������mysql�������˺š����롢���ݿ⡢�˿ڡ�ip���ַ���
    coon = MySQLdb.connect(user='zzf',passwd='123456',db='bugfree',port=3306,host='127.0.0.1',charset='utf8')
    #�����α�
    cursor = coon.cursor()
    #ִ��sql
    cursor.execute(sql)
    #�ύ
    coon.commit()
    #�ر��α�
    cursor.close()
    #�ر�����
    coon.close()


if __name__ == '__main__':
    try:
        filename = 'test_case.xls'
        print filename
        readExcel(filename)
        print 'ִ����ϣ���鿴���Խ����'
        thisIsLove = input('�� ENTER ���˳����ڣ�')
    except Exception,e:
        print e