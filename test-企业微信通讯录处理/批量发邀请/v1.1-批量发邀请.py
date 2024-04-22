# -*- coding: UTF-8 -*-

import pandas as pd
import requests
import json
import logging
import traceback
import time


logging.basicConfig(level=logging.INFO,    # 控制台打印的日志级别
            filename='log.txt',
            filemode='a',   # 模式，有w和a，w就是写模式，默认是追加模式
            format='%(asctime)s - [line:%(lineno)d] - %(levelname)s: %(message)s',  # 日志格式
            encoding='utf8')


def decryption(str_secret,salt):
    list_secret=str_secret.split('O')

    list_salt=salt.split('OO')

    len_prefix=len(str(int(list_salt[0],2)))

    str_indexes=list_salt[1]

    list_index=str_indexes.split('O')
    list_index_10=[int(i,2) for i in list_index]
    list_sort=list_index_10[:len_prefix]
    list_sort.sort()

    for i in range(len_prefix-1,-1,-1):
        index=int(list_sort[i])
        del(list_secret[index])

    list_char=[]

    for i in range(len(list_secret)):
        num=int(list_secret[i],16)+int(list_index_10[i])
        c=chr(num)
        list_char.append(c)

    str_password=''.join(list_char)

    return str_password


def login(ID,SECRET):

    url='https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid=%s&corpsecret=%s' %(ID,SECRET)

    result_return=requests.request('get',url,headers=headers).json()

    # {
    # "errcode": 0,
    # "errmsg": "ok",
    # "access_token": "accesstoken000001",
    # "expires_in": 7200
    # }

    return result_return


def invite(ACCESS_TOKEN,list_user_id):
    '''
    请求方式：POST（HTTPS）
    请求地址： https://qyapi.weixin.qq.com/cgi-bin/batch/invite?access_token=ACCESS_TOKEN

    请求示例：
    {
    "user" : ["UserID1", "UserID2", "UserID3"],
    "party" : [PartyID1, PartyID2],
    "tag" : [TagID1, TagID2]
    }

    参数说明：
    参数            是否必须    说明
    access_token    是          调用接口凭证
    user            否          成员ID列表, 最多支持1000个。
    party           否          部门ID列表，最多支持100个。
    tag             否          标签ID列表，最多支持100个。

    返回示例：
    {
    "errcode" : 0,
    "errmsg" : "ok",
    "invaliduser" : ["UserID1", "UserID2"],
    "invalidparty" : [PartyID1, PartyID2],
    "invalidtag": [TagID1, TagID2]
    }

    参数说明：
    参数            说明
    errcode         返回码
    errmsg          对返回码的文本描述内容
    invaliduser     非法成员列表
    invalidparty    非法部门列表
    invalidtag      非法标签列表

    更多说明：
    user, party, tag三者不能同时为空；
    如果部分接收人无权限或不存在，邀请仍然执行，但会返回无效的部分（即invaliduser或invalidparty或invalidtag）;
    同一用户只须邀请一次，被邀请的用户如果未安装企业微信，在3天内每天会收到一次通知，最多持续3天。
    因为邀请频率是异步检查的，所以调用接口返回成功，并不代表接收者一定能收到邀请消息（可能受上述频率限制无法接收）。
    '''

    url='https://qyapi.weixin.qq.com/cgi-bin/batch/invite?access_token=%s' % ACCESS_TOKEN

    len_userid=len(list_user_id)

    index_next=0
    step=1000
    for i in range(0,len_userid,step):
        index_next=i+step
        if index_next>len_userid:
            index_next=len_userid

        data={ 'user': list_user_id[i:index_next] }

        result_return=requests.request('post',url,data=json.dumps(data),headers=headers).json()

        str_info='批量邀请[%s-%s]/%s：\n\n%s\n\n运行信息：%s' %(i+1,index_next,len_userid,data,result_return)
        logging.info(str_info)
        print(str_info)


headers={'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0'}


def main():
    df=pd.read_excel('要邀请的信息.xlsx',header=0,usecols=['userid'])

    if df.empty: return

    with open('../key','r') as f:
        list_key=f.readlines()

    ID=decryption(list_key[0],list_key[1])
    SECRET=decryption(list_key[2],list_key[3])

    print('登录...')

    result_return=login(ID,SECRET)
    if result_return['errcode'] !=0 :
        str_error='登录错误：%s %s' %(result_return['errcode'],result_return['errmsg'])
        logging.error(str_error)
        print(str_error)
        return

    print('邀请...')

    list_user_id=df['userid'].values.tolist()
    # print(list_user_id)

    invite(result_return['access_token'],list_user_id)

    print('完成。')

    time.sleep(6)


if __name__=='__main__':
    try:
        main()
    except Exception:
        str_error='×错误：%s' % traceback.format_exc()
        logging.error(str_error)
        input(str_error)
