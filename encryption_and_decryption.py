import time
import random

def encryption(str_password,prefix):
    rng_secret=range(len(str_password+prefix))
    list_index=[i for i in rng_secret]
    random.shuffle(list_index)

    postfix='O'.join([str(bin(i)) for i in list_index])

    list_ascii_16=[]
    
    for i in range(len(str_password)):
        num=list_index[i]
        num_16=hex(ord(str_password[i])-num)
        list_ascii_16.append(num_16)

    list_secret=['' for i in rng_secret]

    for i in range(len(prefix)):
        index=list_index[i]
        list_secret[index]=hex(ord(prefix[i]))

    i =0
    for j in rng_secret:
        if list_secret[j]=='':
            list_secret[j]=list_ascii_16[i]
            i+=1

    str_secret='O'.join(list_secret)
    salt =str(bin(int(prefix))) + 'OO' + postfix

    return str_secret,salt


def decryption(str_secret,salt):
    list_secret=str_secret.split('O')

    list_salt=salt.split('OO')

    len_prefix=len(str(int(list_salt[0],2)))

    str_indexes=list_salt[1]

    list_index=str_indexes.split('O')
    list_index_10=[int(i,2) for i in list_index]
    list_sort=list_index_10[:len_prefix]
    list_sort=sorted(list_sort)
    # print(list_sort)

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

if __name__=='__main__':
    filename='k_zhangyu'
    with open(filename,'w',encoding='utf-8') as f:
        f.write('')

    prefix=time.strftime('%Y%m%d%H%M')
    print(prefix)

    print('-----加密1-----')
    str_password='DIKJm3908%￥LWFE`￥#@123!'
    str_secret,salt=encryption(str_password,prefix)
    print('%s\n%s' %(str_secret,salt))

    with open(filename,'a',encoding='utf-8') as f:
        f.write(str_secret + '\n' + salt + '\n')

    print('------解密1------')
    str_password=decryption(str_secret,salt)
    print(str_password)

    print('***************************')

    print('-----加密2------')
    str_secret2,salt2=encryption('w!@#$%^&*()_+-=<>?1234567890jd',prefix)
    print('%s\n%s' %(str_secret2,salt2))

    with open(filename,'a',encoding='utf-8') as f:
        f.write(str_secret2 + '\n' + salt2 + '\n')

    print('-------解密2-----------')
    str_password=decryption(str_secret2,salt2)
    print(str_password)
