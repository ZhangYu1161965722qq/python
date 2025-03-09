import pandas as pd
import time

def main():
    df_src=pd.read_excel('原始信息.xlsx',dtype=str)

    # 填充空白
    df_src.fillna('',inplace=True)

    list_df=df_src.values.tolist()

    # id的列表
    list_id=df_src.iloc[:,0].tolist()

    len_list=len(list_df)

    dict_id_level={}
    dict_level_id={}
    for i in range(len_list):
        id=list_df[i][0]
        parent_id=list_df[i][2]

        num_level=0

        # 记录id对应level中间过程的字典
        _dict={}

        # 循环获取向上到根节点的所有负数步骤
        while True:
            # 临时记录差的层次夫负步骤
            num_level -=1
            if num_level not in _dict:
                _dict[num_level]=id

            # 直到父做id时不在id列表中，所有的id都要记录，所以条件不直接写在while上
            if id not in list_id: break

            # 根据id找对应的父id
            parent_id=df_src[df_src.iloc[:,0]==id].iloc[0,2]

            # 拿父ID继续向上循环找
            id=parent_id


        for k,v in _dict.items():
            # 把负数步骤转换成level数
            _level=k-num_level

            # 放入id为键，level为值的字典
            if v not in dict_id_level:
                dict_id_level[v]=_level

            # 放入level为键，id为值的字典
            if _level not in dict_level_id:
                dict_level_id[_level]=[v]
            else:
                _li_id=dict_level_id[_level]
                # 没有放入过的id值就放入
                if v not in _li_id:
                    _li_id.append(v)


    list_result=[]

    # level最大数的位数
    len_level=len(str(len(dict_level_id)))

    for i in range(len_list):
        # 初始id、parent_id
        id=list_df[i][0]
        parent_id=list_df[i][2]

        # 中间循环需要的_id、_parent_id,初始为当前id、parent_id
        _id=id
        _parent_id=parent_id
        
        _level=-1

        # 循环拼接出:level1-顺序号;level2-顺序号;...
        level_seq=''
        while _id in list_id :
            _level=dict_id_level[_id]
            _li=dict_level_id[_level]

            # 顺序号最大数的位数
            _lenB=len(str(len(_li)))

            # 拼接：level-顺序号，加前导0
            level_seq=str(_level).zfill(len_level)+'-' +str(_li.index(_id)).zfill(_lenB)+';'+level_seq

            # 根据id找对应的父id
            _parent_id=df_src[df_src.iloc[:,0]==_id].iloc[0,2]

            # 拿父ID继续向上循环找，直到父id不在id列表中
            _id=_parent_id

        level=dict_id_level[id]

        # 拼接显示名
        name_node=df_src[df_src.iloc[:,0]==id].iloc[0,1]
        name_node= '  ' * level + '|_ ' + name_node

        list_result.append([level_seq,level,id,name_node,parent_id])

    # 排序
    # list_result=sorted(list_result,key=(lambda x:x[0]))
    
    # list转换成dataframe
    df_result=pd.DataFrame(list_result)

    # 排序，按第一列升序
    df_result.sort_values(by=[0],inplace=True,ascending=True)

    # with open('r_id_level.txt','w',encoding='utf8') as f:
    #     f.write(str(dict_id_level))

    # with open('r_level_id.txt','w',encoding='utf8') as f:
    #     f.write(str(dict_level_id))

    # with open('list_result.txt','w',encoding='utf8') as f:
    #     f.write(str(list_result)) 

    # 改变列名
    df_result.columns=['level_seq','level']+df_src.columns.tolist()

    df_result.to_excel('结果-层次.xlsx',index=False)

if __name__=='__main__':
    print(time.strftime('%H:%M:%S'),'Start')
    main()
    print('%s End' %time.strftime('%H:%M:%S'))