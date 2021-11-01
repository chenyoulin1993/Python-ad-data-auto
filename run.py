import pandas as pd
import numpy as np

def sp(x):
    a = str(x).split('_')[0]
    return a

def cp(x):
    a = str(x)[0:2]
    return a

def ri(x):
    a = str(x)[2:8]
    return a
    # return a

def biandao(x):
    a = str(x)[8:9]
    n = bd[bd['编号']==a]['编导'].values[0]
    return n

def jianji(x):
    a = str(x)[9:10]
    n = jj[jj['编号'] == a]['剪辑'].values[0]
    return n

def guoch(over):
    over = pd.DataFrame(over)

    over['千次展现成本'] = over['消耗'] / over['展现量']
    over['点击率'] = over['点击量'] / over['展现量']
    over['点击单价'] = over['消耗'] / over['点击量']
    over['转化率'] = over['转化数'] / over['点击量']
    over['转化成本'] = over['消耗'] / over['转化数']
    over['加购成本'] = over['消耗'] / over['添加购物车量']
    over['投资回报率'] = over['成交订单金额'] / over['消耗']
    over = over.round(2)
    over = over.replace(np.inf, 0)

    over['产品'] = over['素材'].map(cp)
    over['交付日期'] = over['素材'].map(ri)
    over['编导'] = over['素材'].map(biandao)
    over['剪辑'] = over['素材'].map(jianji)

    over = over[['日期', '素材', '消耗', '展现量', '点击量', '转化数', '成交订单量',
                 '成交订单金额', '添加购物车量', '千次展现成本'
        , '点击率', '点击单价', '转化率', '转化成本', '加购成本', '投资回报率', '产品', '交付日期', '编导', '剪辑'
                 ]]
    return over

def sucaijiaofu(df):

    df = df.reset_index(drop=True)
    df['数量']=1

    jj = df.groupby(['日期','剪辑']).sum()

    jj['投资回报率'] = jj['成交订单金额'] / jj['消耗']
    jj['点击率'] = jj['点击量'] / jj['展现量']

    jj = jj[['数量','消耗','成交订单金额','投资回报率','展现量', '点击量','点击率']]


    bd = df.groupby(['日期','编导']).sum()
    bd['投资回报率'] = bd['成交订单金额'] / bd['消耗']
    bd['点击率'] = bd['点击量'] / bd['展现量']

    bd = bd[['数量', '消耗', '成交订单金额', '投资回报率', '展现量', '点击量', '点击率']]
    jj = jj.round(2)
    bd = bd.round(2)

    return jj,bd



if __name__ == '__main__':
    df = pd.read_excel('UD数据表.xls')
    bd = pd.read_excel('素材底表.xlsx',sheet_name='编导')
    jj = pd.read_excel('素材底表.xlsx',sheet_name='剪辑')

    df = pd.DataFrame(df)
    df = df.sort_values('日期')
    df['素材'] = df['计划'].map(sp)
    ba = pd.read_excel('素材底表.xlsx',sheet_name='素材')
    ba_list = ba['素材'].values
    df1 = df[df['素材'].isin(ba_list)]
    df2 = df[~df['素材'].isin(ba_list)]
    over = df1.groupby(['日期','素材']).sum()[['消耗', '展现量', '点击量', '转化数', '成交订单量', '成交订单金额', '添加购物车量']]

    over['日期'] = over.index.get_level_values(0).tolist()
    over['素材'] = over.index.get_level_values(1).tolist()
    over = guoch(over)


    week_df = over.copy()
    week_df = week_df.droplevel(level=1)
    week_df['日期'] = pd.to_datetime(week_df['日期'])
    week_df.index = week_df['日期']

    sucai = week_df['素材'].values.tolist()
    sucai = list(set(sucai))

    week_over = pd.DataFrame()
    month_over = pd.DataFrame()
    for i in sucai:
        tes = week_df[week_df['素材']==i]
        week = tes.resample('W').sum()
        month = tes.resample('M').sum()
        week['素材'] = i
        month['素材'] = i

        week_over = week_over.append(week)
        month_over = month_over.append(month)

    week_over['日期'] = week_over.index
    month_over['日期'] = month_over.index


    week_over = guoch(week_over)
    month_over = guoch(month_over)

    week_copy = week_over.copy()
    month_copy = month_over.copy()

    jj1,bd1 = sucaijiaofu(week_copy)
    jj2,bd2 = sucaijiaofu(month_copy)

    week_sucai = jj1.append(bd1)
    month_sucai = jj2.append(bd2)
    # print(jj)
    # print(bd)


    zong = df1.groupby(['素材']).sum()[['消耗', '展现量', '点击量', '转化数', '成交订单量', '成交订单金额', '添加购物车量']]
    zong['素材'] = zong.index

    zong['千次展现成本'] = zong['消耗'] / zong['展现量']
    zong['点击率'] = zong['点击量'] / zong['展现量']
    zong['点击单价'] = zong['消耗'] / zong['点击量']
    zong['转化率'] = zong['转化数'] / zong['点击量']
    zong['转化成本'] = zong['消耗'] / zong['转化数']
    zong['加购成本'] = zong['消耗'] / zong['添加购物车量']
    zong['投资回报率'] = zong['成交订单金额'] / zong['消耗']
    zong = zong.round(2)
    zong = zong.replace(np.inf, 0)

    zong['产品'] = zong['素材'].map(cp)
    zong['交付日期'] = zong['素材'].map(ri)
    zong['编导'] = zong['素材'].map(biandao)
    zong['剪辑'] = zong['素材'].map(jianji)

    zong = zong[[ '素材', '消耗', '展现量', '点击量', '转化数', '成交订单量',
                 '成交订单金额', '添加购物车量', '千次展现成本'
        , '点击率', '点击单价', '转化率', '转化成本', '加购成本', '投资回报率', '产品', '交付日期', '编导', '剪辑'
                 ]]



    with pd.ExcelWriter(r'ud结果表.xlsx') as writer:
        over.to_excel(writer, sheet_name='日',index=False)
        week_over.to_excel(writer, sheet_name='周',index=False)
        month_over.to_excel(writer, sheet_name='月',index=False)
        zong.to_excel(writer, sheet_name='总',index=False)
        week_sucai.to_excel(writer, sheet_name='周素材汇总')
        month_sucai.to_excel(writer, sheet_name='月素材汇总')


