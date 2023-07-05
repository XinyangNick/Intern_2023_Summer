#!/usr/bin/env python
# coding: utf-8

# # 输入文件名

# In[1]:


import pandas as pd

selction = int(input('请选择模式1生成结果2生成模板:'))
# In[ ]:


if selction == 1:

    Selling_name = str(input('请输入销售发票申请文件名:')) + '.xlsx'


    # In[ ]:


    Invoice_name = str(input('请输入收票核销记录文件名：')) + '.xlsx'


    # In[ ]:


    transfer_name = str(input('请输入采销关联表文件名：')) + '.xlsx'


    # In[ ]:


    Selling_order_name = str(input('请输入销售订单文件名：')) + '.xlsx'


    # In[ ]:


    Stardard_name = str(input('请输入国税全量发票查询文件名：')) + '.xlsx'


    # In[ ]:


    Buyer_name = str(input('请输入客户信息文件名：')) + '.xlsx'


    # # 测试名字

    # In[319]:


    # Selling_name = '销售发票申请2023-07-04.xlsx'
    # Invoice_name = '收票核销记录2023-07-04.xlsx'
    # Selling_order_name = '销售订单2023-07-04.xlsx'
    # Stardard_name = '全量发票查询导出结果.xlsx'
    # transfer_name = '采销关联表2023-07-04.xlsx'
    # Buyer_name = '客户信息20230704a1f47641.xlsx'


    # # 读取文件

    # In[320]:


    Selling_df = pd.read_excel(Selling_name)
    Invoice_df = pd.read_excel(Invoice_name)
    Selling_order_df = pd.read_excel(Selling_order_name)
    Stardard_df = pd.read_excel(Stardard_name, sheet_name='信息汇总表', dtype=str)
    Stardard_df = Stardard_df.sort_values(by='开票日期')
    transfer_df = pd.read_excel(transfer_name)


    # In[321]:


    client_df = Selling_df.drop_duplicates(subset=['客户'])['客户']
    client_df.to_excel('客户列表.xlsx')


    # In[ ]:


    Buyer_df = pd.read_excel(Buyer_name)


    # In[322]:


    Invoice_df_group = Invoice_df.groupby('采购订单号').sum()
    Selling_order_df_group = Selling_order_df.groupby('销售订单号').sum()


    # In[323]:


    #进项重量 和 销售订单金额
    amout_dict = dict(zip(Invoice_df_group.index, Invoice_df_group['发票重量'])) #采购订单号：发票重量


    # In[324]:


    #销售找采购号
    transfer_dict = {}
    for sell_code in transfer_df['销售订单号']:
        transfer_dict[sell_code] = set(transfer_df.loc[transfer_df['销售订单号'] == sell_code]['采购订单号'])


    # In[325]:


    Invoice_buying_code = list(Invoice_df['采购订单号'])
    debt_compony = list(Selling_order_df['销售订单号'])


    # In[326]:


    #匹配采购订单号
    Selling_df['采购订单号'] = ''
    for rows in Selling_df.iterrows():
        index, columns = rows
        selling_code = columns['销售订单号']

        if selling_code in transfer_dict.keys():
            temp_lst = ''
            for buy_code in transfer_dict[selling_code]:
                if buy_code in Invoice_buying_code: #并且采购订单号在进项发票中
                    temp_lst += (buy_code + ',')
            Selling_df.at[index, '采购订单号'] = temp_lst


    # In[327]:


    #复制未开票重量
    Selling_df['已开票重量2'] = Selling_df['已开票重量']
    Selling_df['未开票重量2'] = Selling_df['未开票重量']

    Selling_df['应开重量'] = Selling_df['未开票重量'] - Selling_df['已开票重量']


    # In[328]:


    #销售订单号：所有采购订单号总重量
    sell_amout_dict = {}
    for sell_code in transfer_dict.keys():
        x = 0
        for buy_code in transfer_dict[sell_code]:
            if buy_code in Invoice_buying_code: #并且采购订单号在进项发票中
                x += amout_dict[buy_code]
        sell_amout_dict[sell_code] = x


    # In[329]:


    #填入总重量
    for rows in Selling_df.iterrows():
        index, columns = rows

        sell_code = columns['销售订单号']

        if sell_code in sell_amout_dict.keys():
            Selling_df.at[index, '进项总重量'] = sell_amout_dict[sell_code]


    # In[330]:


    #计算应开重量
    calulate_amount_dict = sell_amout_dict.copy()

    for rows in Selling_df.iterrows():
        index, columns = rows

        sell_code = columns['销售订单号']
        apply_amount = columns['应开重量']
        total_amount = columns['进项总重量']
        remain_amount = calulate_amount_dict[sell_code]

        if remain_amount > 0:
            x = min(remain_amount, apply_amount)

            Selling_df.at[index, '可开重量'] = x
            calulate_amount_dict[sell_code] -= x
            Selling_df.at[index, '进项剩余重量'] = calulate_amount_dict[sell_code]



    # In[331]:


    #匹配金额
    for rows in Selling_df.iterrows():
        index, columns = rows

        sell_code = columns['销售订单号']
        if sell_code in Selling_order_df_group.index:
            Selling_df.at[index, '应收金额'] = Selling_order_df_group.loc[sell_code]['应收金额']
            Selling_df.at[index, '已收金额'] = Selling_order_df_group.loc[sell_code]['已收金额']


    # In[332]:


    #判断可不可开
    for rows in Selling_df.iterrows():
        index, columns = rows

        department = columns['部门']
        should_pay = columns['应开重量']
        able_pay = columns['可开重量']

        should_recive_money = columns['应收金额']
        received_money = columns['已收金额']

        pay_state = None
        if department == '业务二部': #如果是业务二部 不需要管金额
            if able_pay > 0 and able_pay == should_pay:
                pay_state = '可全开'
            elif able_pay > 0 and able_pay != should_pay:
                pay_state = '可部分开'
            else:
                pay_state = '不可开'
        else: #除了业务二部的
            if received_money == should_recive_money and should_pay == able_pay:
                pay_state = '可全开'
            else:
                pay_state = '不可开'

        Selling_df.at[index, '可开状态'] = pay_state

        #尝试匹配供应商
        buy_code_lst = columns['采购订单号']
        temp_index = Invoice_df.set_index('采购订单号')
        temp_name = ''
        temp_name_lst = []

        for buy_codes in buy_code_lst.split(','):
            if buy_codes in temp_index.index and buy_codes != '':
                x = temp_index.loc[buy_codes]['开票单位']
                if x not in temp_name_lst:
                    temp_name_lst.append(list(x)[0])
        for item in set(temp_name_lst):
            temp_name += (item + ',')

        Selling_df.at[index, '供应商'] = temp_name


    # In[333]:


    Selling_df_match = Selling_df.copy()
    Selling_df_match['品名匹配率'] = 0
    Selling_df_match['牌号匹配率'] = 0


    # In[334]:


    def spec_ratio(good_spec: str, good_speci2: str):
        count = good_spec
        for elem in good_spec:
            if elem in good_speci2:
                count = count[1:]
        return (len(good_spec) - len(count))/ len(good_spec) * 100

    def spec_ratio2(good_spec: str, good_spec2: str):
        return (spec_ratio(good_spec, good_spec2) + spec_ratio(good_spec2, good_spec)) / 2


    # In[335]:


    Selling_df_match['品名2'] = Selling_df_match['品名']
    Selling_df_match['牌号2'] = Selling_df_match['牌号']


    # In[336]:


    for rows in Selling_df_match.iterrows():
        index, columns = rows

        good_name = columns['品名']
        good_spec = columns['牌号']
        good_name_match_rate = columns['品名匹配率']
        good_spec_match_rate = columns['牌号匹配率']

        supply_name_lst = columns['供应商'].split(',')
        for supply_name in supply_name_lst:
            if supply_name != '':

                if isinstance(supply_name, str):
                    high_good_name = None
                    high_good_spec = None
                    high_good_name_match_rate = good_name_match_rate
                    high_good_spec_match_rate = good_spec_match_rate
                    high_standard_good_code = 0

                    temp_df = Stardard_df.loc[Stardard_df['销方名称'] == supply_name]

                    possible_list = []
                    for sub_rows in temp_df.iterrows():
                        sub_index, sub_columns = sub_rows
                        if isinstance(sub_columns['货物或应税劳务名称'], str) and '*' in sub_columns['货物或应税劳务名称']:
                            standard_good_name = sub_columns['货物或应税劳务名称'].split('*')[2]
                        standard_good_spec = str(sub_columns['规格型号'])
                        standard_good_code = str(sub_columns['税收分类编码'])

                        func_name_match_rate = spec_ratio2(good_name, standard_good_name)
                        func_spec_match_rate = spec_ratio2(good_spec, standard_good_spec)



                        if str(good_spec) in str(standard_good_spec):
                            high_good_name = standard_good_name
                            high_good_spec = standard_good_spec
                            high_good_name_match_rate = func_name_match_rate
                            high_good_spec_match_rate = func_spec_match_rate
                            high_standard_good_code = standard_good_code

                        elif func_name_match_rate >= high_good_name_match_rate and func_spec_match_rate > high_good_spec_match_rate:
                            high_good_name = standard_good_name
                            high_good_spec = standard_good_spec
                            high_good_name_match_rate = func_name_match_rate
                            high_good_spec_match_rate = func_spec_match_rate
                            high_standard_good_code = standard_good_code

                    Selling_df_match.at[index, '标准品名'] = high_good_name
                    Selling_df_match.at[index, '标准牌号'] = high_good_spec
                    Selling_df_match.at[index, '品名匹配率'] = high_good_name_match_rate
                    Selling_df_match.at[index, '牌号匹配率'] = high_good_spec_match_rate
                    Selling_df_match.at[index, '税收分类编码'] = high_standard_good_code


                    if high_good_name_match_rate < 50 or high_good_spec_match_rate < 50:
                        Selling_df_match.at[index, '潜在问题'] = '匹配率过低'


    # In[337]:


    Selling_df_match['含税单价'] = Selling_df_match['未开票金额'] / Selling_df_match['未开票重量']


    # In[338]:


    Selling_df_match['应开金额3'] = Selling_df_match['含税单价'] *  Selling_df_match['可开重量']


    # # 发票模板

    # In[339]:


    Buyer_dict = dict(zip(Buyer_df['客户名称'], Buyer_df['统一社会信用代码/纳税人识别号']))


    # In[340]:


    pay_selling_df = Selling_df_match.loc[Selling_df_match['可开状态'] != '不可开']


    # In[341]:


    for rows in pay_selling_df.iterrows():
        index, columns = rows
        try:
            pay_selling_df.at[index, '统一社会信用代码/纳税人识别号'] = Buyer_dict[columns['客户']]
        except:
            pay_selling_df.at[index, '统一社会信用代码/纳税人识别号'] = 0


    # In[342]:


    for rows in Selling_df_match.iterrows():
        index, columns = rows
        try:
            Selling_df_match.at[index, '统一社会信用代码/纳税人识别号'] = Buyer_dict[columns['客户']]
        except:
            Selling_df_match.at[index, '统一社会信用代码/纳税人识别号'] = 0


    # In[343]:


    Selling_df_match.to_excel('0-完整未开票（结果）.xlsx')


    # In[344]:


    pay_selling_df.to_excel('0-只含需要开票文件(结果).xlsx')


# # 发票模板

# In[345]:
elif selction == 2:

    pay_selling_df = pd.read_excel('0-只含需要开票文件(结果).xlsx', dtype=str)


    # In[346]:


    result_basic = pd.DataFrame([['发票流水号', '发票类型', '特定业务类型', '是否含税', '受票方自然人标识', '购买方名称', '证件类型',
           '购买方纳税人识别号', '购买方地址', '购买方电话', '购买方开户银行', '购买方银行账号', '备注',
           '是否展示购买方银行账号', '销售方开户行', '销售方银行账号', '是否展示销售方银行账号', '购买方邮箱', '购买方经办人姓名',
           '购买方经办人证件类型', '购买方经办人证件号码', '经办人国籍(地区)', '经办人自然人纳税人识别号',
           '放弃享受减按1%征收率\n原因', '收款人', '复核人']])


    # In[347]:


    result_detail = pd.DataFrame([['发票流水号', '项目名称', '商品和服务税收编码', '规格型号', '单位', '数量', '单价', '金额', '税率',
           '折扣金额', '是否使用优惠政策', '优惠政策类型', '即征即退类型', 'index']])


    # In[348]:


    for rows in pay_selling_df.iterrows():
        index, value = rows
        temp = pd.DataFrame([[value['销售订单号'], '增值税专用发票', '',
                              '是',
                              '',
                              value['客户'],
                              '',
                              value['统一社会信用代码/纳税人识别号'],
                              '',
                              '',
                              '',
                              '',
                              ''
                              ]])
        result_basic = pd.concat([result_basic, temp])


    # In[349]:


    result_basic = result_basic.drop_duplicates(subset=0)


    # In[350]:


    result_basic.to_excel('1-发票基本信息.xlsx')


    # In[351]:


    for rows in pay_selling_df.iterrows():
        index, value = rows
        temp = pd.DataFrame([[value['销售订单号'], value['标准品名'], value['税收分类编码'], value['标准牌号'],
                              '吨', value['可开重量'], value['含税单价'], value['应开金额3'], 0.13]])
        result_detail = pd.concat([result_detail, temp])


    # In[352]:


    result_detail.to_excel('2-发票明细信息.xlsx')

