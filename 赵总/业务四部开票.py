#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd

selection = int(input('请选择模式输入1生成结果2生成模板:'))
# In[ ]:
if selection == 1:

    sell_name = str(input('请输入销售未开票文件名：')) + '.xlsx'


    # In[2]:


    DaoEn_df = pd.read_excel(sell_name)


    # In[ ]:


    invoice_name = str(input('请输入售票核销记录文件名：')) + '.xlsx'


    # In[3]:


    Invoice_df = pd.read_excel(invoice_name)


    # In[4]:


    DaoEn_df['未开票重量+已开票重量'] = DaoEn_df['未开票重量'] + DaoEn_df['已开票重量']


    # In[5]:


    DaoEn_group = DaoEn_df.groupby('入库批号').sum()
    Invoice_group = Invoice_df.groupby('入库批号').sum()


    # In[6]:


    DaoEn_group_jin = DaoEn_group.copy()
    DaoEn_group_jin['未开票重量2'] = DaoEn_group_jin['未开票重量']
    DaoEn_group_jin['已开票重量2'] = DaoEn_group_jin['已开票重量']


    for rows in DaoEn_group_jin.iterrows():
        index_value, column_value = rows
        if index_value in Invoice_group.index:
            DaoEn_group_jin.at[index_value, '进项重量'] = Invoice_group.loc[index_value]['发票重量']
        else:
            DaoEn_group_jin.at[index_value, '进项重量'] = 0


    # In[7]:


    DaoEn_group_jin['最大可开数量'] = DaoEn_group_jin['进项重量'] - DaoEn_group_jin['已开票重量']


    # In[8]:


    for rows in DaoEn_group_jin.iterrows():
        index_value, column_value = rows
        x = ''
        if column_value['最大可开数量'] == column_value['未开票重量']:
            x = '可全开'
        elif column_value['最大可开数量'] > 0 and column_value['最大可开数量'] < column_value['未开票重量']:
            x = '可以部分开'
        elif  column_value['最大可开数量'] == 0:
            x = '不可以开'
        elif column_value['最大可开数量'] > column_value['未开票重量']:
            x = '可全开'
        DaoEn_group_jin.at[index_value, '可开票状态'] = x
        DaoEn_group_jin.at[index_value, '应开数量'] = min(column_value['最大可开数量'], column_value['未开票重量'])


    # In[9]:


    col_lst = DaoEn_group_jin.columns[-7:]


    # In[10]:


    lst_code = list(DaoEn_df['入库批号'])
    for rows in DaoEn_df.iterrows():
        index_value, column_value = rows
        cout_num = lst_code.count(column_value['入库批号'])
        if cout_num > 1:
            DaoEn_df.at[index_value, '入库批号是否重复'] = '是'
        else:
            DaoEn_df.at[index_value, '入库批号是否重复'] = '否'
        for col in col_lst:
            DaoEn_df.at[index_value, col] = DaoEn_group_jin.loc[column_value['入库批号']][col]



    # In[11]:


    DaoEn_df['含税单价2'] = DaoEn_df['含税单价']


    # In[12]:


    DaoEn_df['应开金额'] = DaoEn_df['含税单价'] * DaoEn_df['应开数量']


    # In[13]:


    neg_lst = []
    for rows in DaoEn_df.iterrows():
        index_value, column_value = rows
        if column_value['无税金额'] <= 0:
            neg_lst.append(column_value['销售订单号'])



    # In[14]:


    for rows in DaoEn_df.iterrows():
        index_value, column_value = rows
        if column_value['销售订单号'] in neg_lst:
            DaoEn_df.at[index_value, '相同销售订单号含负数金额'] = '是'
        else:
            DaoEn_df.at[index_value, '相同销售订单号含负数金额'] = '否'


    # In[15]:


    Invoice_df_index = Invoice_df.drop_duplicates('入库批号').set_index('入库批号')


    # In[16]:


    for rows in DaoEn_df.iterrows():
        index_value, column_value = rows
        pihao = column_value['入库批号']
        if pihao in Invoice_df_index.index and column_value['应开数量'] > 0:
            DaoEn_df.at[index_value, '供应商名字'] = Invoice_df_index.loc[pihao]['开票单位']
        else:
            DaoEn_df.at[index_value, '供应商名字'] = None


    # In[ ]:


    stardard_name = str(input('请输入全量发票导出查询结果：')) + '.xlsx'


    # In[17]:


    standard_invoice = pd.read_excel(stardard_name, sheet_name='信息汇总表', dtype={'发票代码': str,
                                                                                      '发票号码': str,'数电票号码': str,
                                                                                     '税收分类编码': str})
    standard_invoice = standard_invoice.sort_values(by='开票日期')


    # In[18]:


    standard_invoice['销售单价'] = standard_invoice['单价'] * 1.13 + 10


    # In[19]:


    def spec_ratio(good_spec: str, good_speci2: str):
        count = good_spec
        for elem in good_spec:
            if elem in good_speci2:
                count = count[1:]
        return (len(good_spec) - len(count))/ len(good_spec) * 100

    def spec_ratio2(good_spec: str, good_spec2: str):
        return (spec_ratio(good_spec, good_spec2) + spec_ratio(good_spec2, good_spec)) /2


    # In[20]:


    DaoEn_df['牌号匹配率'] = 0
    DaoEn_df['品名匹配率'] = 0

    for rows in DaoEn_df.iterrows():
        index_value, column_value = rows
        supply_name = column_value['供应商名字']
        taxed_price = column_value['含税单价']
        good_type = column_value['品名']
        good_spec = column_value['牌号']
        match_rate = column_value['牌号匹配率']
        type_match_rate = column_value['品名匹配率']

        if supply_name != None:
            temp_df = standard_invoice.loc[standard_invoice['销方名称'] == supply_name].copy()

            high_good_name = None
            high_good_code = None
            high_good_spec = None
            high_fun_match = match_rate
            high_type_match_rate = type_match_rate

            for sub_rows in temp_df.iterrows():
                sub_index_value, sub_column_value = sub_rows
                good_code = sub_column_value['税收分类编码']
                good_name = sub_column_value['货物或应税劳务名称'].split('*')[2]
                good_specification = sub_column_value['规格型号']
                selling_price = sub_column_value['销售单价']

                fun_match = spec_ratio2(good_spec, good_specification)
                fun_type_match = spec_ratio2(good_type, good_name)

                if abs(selling_price - taxed_price) < 0.05:

                    if fun_type_match >= high_type_match_rate and fun_match > high_fun_match:
                        high_good_name = good_name
                        high_good_spec = good_specification

                        high_fun_match = fun_match
                        high_type_match_rate = fun_type_match

                        high_good_code = good_code


            DaoEn_df.at[index_value, '货物或应税劳务名称'] = high_good_name
            DaoEn_df.at[index_value, '规格型号'] = high_good_spec

            DaoEn_df.at[index_value, '牌号匹配率'] = high_fun_match
            DaoEn_df.at[index_value, '品名匹配率'] = high_type_match_rate

            DaoEn_df.at[index_value, '税收分类编码'] = high_good_code





    # In[ ]:


    # compony_code_name = str(input('请输入客户开票资料文件名：')) + '.xlsx'
    #
    #
    # # In[21]:
    #
    #
    # compony_code = pd.read_excel(compony_code_name, header=None, names=['compony', 'code']).set_index('compony')


    # In[22]:


    DaoEn_df.to_excel('未开票（结果）.xlsx')



elif selection == 2:

# In[23]:
    compony_code_name = str(input('请输入客户开票资料文件名：')) + '.xlsx'


    # In[21]:


    compony_code = pd.read_excel(compony_code_name, header=None, names=['compony', 'code']).set_index('compony')

    ready_df = pd.read_excel('未开票（结果）.xlsx', dtype=str)


    # In[24]:


    ready_df = ready_df.dropna(subset=['供应商名字'])


    # In[25]:


    for rows in ready_df.iterrows():
        index_value, column_value = rows
        compony_name = column_value['客户']
        ready_df.at[index_value, '购买方纳税人识别号'] = compony_code.loc[compony_name]['code']


    # In[26]:


    result_basic = pd.DataFrame([['发票流水号', '发票类型', '特定业务类型', '是否含税', '受票方自然人标识', '购买方名称', '证件类型',
           '购买方纳税人识别号', '购买方地址', '购买方电话', '购买方开户银行', '购买方银行账号', '备注',
           '是否展示购买方银行账号', '销售方开户行', '销售方银行账号', '是否展示销售方银行账号', '购买方邮箱', '购买方经办人姓名',
           '购买方经办人证件类型', '购买方经办人证件号码', '经办人国籍(地区)', '经办人自然人纳税人识别号',
           '放弃享受减按1%征收率\n原因', '收款人', '复核人']])


    # In[27]:


    result_detail = pd.DataFrame([['发票流水号', '项目名称', '商品和服务税收编码', '规格型号', '单位', '数量', '单价', '金额', '税率',
           '折扣金额', '是否使用优惠政策', '优惠政策类型', '即征即退类型', 'index']])


    # In[28]:


    for rows in ready_df.iterrows():
        index, value = rows
        temp = pd.DataFrame([[value['销售订单号'], '增值税专用发票', '',
                              '是',
                              '',
                              value['客户'],
                              '',
                              value['购买方纳税人识别号'],
                              '',
                              '',
                              '',
                              '',
                              value['销售订单号']
                              ]])
        result_basic = pd.concat([result_basic, temp])


    # In[29]:


    result_basic = result_basic.drop_duplicates(subset=0)


    # In[30]:


    result_basic.to_excel('1-发票基本信息.xlsx')


    # In[31]:


    for rows in ready_df.iterrows():
        index, value = rows
        temp = pd.DataFrame([[value['销售订单号'], value['货物或应税劳务名称'], value['税收分类编码'], value['规格型号'],
                              '吨', value['应开数量'], value['含税单价'], value['应开金额'], 0.13]])
        result_detail = pd.concat([result_detail, temp])


    # In[32]:


    result_detail.to_excel('2-发票明细信息.xlsx')

