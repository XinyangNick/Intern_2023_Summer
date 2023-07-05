#!/usr/bin/env python
# coding: utf-8

# # 输入文件名

# In[1]:


import pandas as pd

selection = int(input('请选择模式输入1生成结果2生成模板:'))

# In[2]:
if selection == 1:

    Selling_name = str(input('请输入销售未开票文件名:')) + '.xlsx'


    # In[3]:


    Invoice_name = str(input('请输入收票核销记录文件名：')) + '.xlsx'


    # In[4]:


    Late_name = str(input('请输入销售逾期报表文件名：')) + '.xlsx'


    # In[5]:


    Stardard_name = str(input('请输入国税全量发票查询文件名：')) + '.xlsx'


    # In[6]:


    Buyer_name = str(input('请输入客户信息文件名：')) + '.xlsx'


    # # 转换Dataframe

    # In[7]:


    Selling_df = pd.read_excel(Selling_name)


    # In[8]:


    Invoice_df = pd.read_excel(Invoice_name)


    # In[9]:


    Late_df = pd.read_excel(Late_name)


    # In[10]:


    Stardard_df = pd.read_excel(Stardard_name, sheet_name='信息汇总表', dtype=str)
    Stardard_df = Stardard_df.sort_values(by='开票日期')


    # In[11]:


    Buyer_df = pd.read_excel(Buyer_name)


    # # 开始操作

    # In[12]:


    Selling_df = Selling_df.loc[Selling_df['部门'] != '业务四部']


    # In[13]:


    Selling_df = Selling_df.loc[Selling_df['部门'] != '业务一部']


    # In[14]:


    Invoice_df_group = Invoice_df.groupby('入库批号').sum()


    # In[15]:


    debt_compony = list(Late_df['实提单号'])
    warehouse_code = list(Selling_df['入库批号'])
    amout_dict = dict(zip(Invoice_df_group.index, Invoice_df_group['发票重量']))
    Invoice_warehouse_code = list(Invoice_df['入库批号'])
    Selling_df['已开票申请重量2'] = Selling_df['已开票申请重量']
    total_amout_dict = amout_dict.copy()

    for rows in Selling_df.iterrows():
        index, columns = rows
        single_warehouse_code = columns['入库批号']
        single_check_code = columns['单据号']
        applied_amout = columns['已开票申请重量']

        try:
            if single_check_code in debt_compony:
                Selling_df.at[index, '是否有欠款'] = '是'
            else:
                Selling_df.at[index, '是否有欠款'] = '否'

            if warehouse_code.count(single_warehouse_code) > 1:
                Selling_df.at[index, '该入库批号是否重复'] = '是'
            else:
                Selling_df.at[index, '该入库批号是否重复'] = '否'


            if single_warehouse_code in Invoice_warehouse_code and len(single_warehouse_code) > 0:
                remain_amout = amout_dict[single_warehouse_code]
                Selling_df.at[index, '进项总重量'] = total_amout_dict[single_warehouse_code]
                if applied_amout <= remain_amout:
                    Selling_df.at[index, '应开重量'] = applied_amout
                    amout_dict[single_warehouse_code] -= applied_amout
        except:
            pass

    # In[16]:


    for rows in Selling_df.iterrows():
        index, columns = rows
        apply_amout = columns['已开票申请重量']
        should_pay_amout = columns['应开重量']
        if apply_amout == should_pay_amout and apply_amout > 0 and columns['是否有欠款'] == '否':
            Selling_df.at[index, '应开金额'] = columns['已开票申请金额']
            Selling_df.at[index, '是否要开'] = '是'
        else:
            Selling_df.at[index, '是否要开'] = '否'


    # In[17]:


    indexd_invoice_df = Invoice_df.drop_duplicates(subset=['入库批号']).set_index('入库批号')
    for rows in Selling_df.iterrows():
        index, columns = rows
        pay_state = columns['是否要开']
        single_warehouse_code = columns['入库批号']
        if pay_state == '是':
            Selling_df.at[index, '供应商'] = indexd_invoice_df.loc[single_warehouse_code]['供应商']
            Selling_df.at[index, '品名2'] = columns['品名']
            Selling_df.at[index, '牌号2'] = columns['牌号']


    # # 匹配国税品名牌号

    # In[18]:


    Selling_df_match = Selling_df.copy()


    # In[19]:


    Selling_df_match['品名匹配率'] = 0
    Selling_df_match['牌号匹配率'] = 0


    # In[20]:


    def spec_ratio(good_spec: str, good_speci2: str):
        count = good_spec
        for elem in good_spec:
            if elem in good_speci2:
                count = count[1:]
        return (len(good_spec) - len(count))/ len(good_spec) * 100



    # In[21]:


    def spec_ratio2(good_spec: str, good_spec2: str):
        return (spec_ratio(good_spec, good_spec2) + spec_ratio(good_spec2, good_spec)) /2


    # In[22]:


    for rows in Selling_df_match.iterrows():
        index, columns = rows
        supply_name = columns['供应商']
        good_name = columns['品名']
        good_spec = columns['牌号']
        good_name_match_rate = columns['品名匹配率']
        good_spec_match_rate = columns['牌号匹配率']


        try:
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

                    possible_good_name = None
                    possible_good_spec = None
                    possible_good_code = None

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

                    elif func_name_match_rate >= high_good_name_match_rate and func_spec_match_rate == high_good_spec_match_rate:
                        possible_good_name = standard_good_name
                        possible_good_spec = standard_good_spec
                        possible_good_code = standard_good_code
                        temp_tuple = [standard_good_name, standard_good_spec, standard_good_code]
                        if not temp_tuple in possible_list:
                            possible_list.append(temp_tuple)


                Selling_df_match.at[index, '标准品名'] = high_good_name
                Selling_df_match.at[index, '标准牌号'] = high_good_spec
                Selling_df_match.at[index, '品名匹配率'] = high_good_name_match_rate
                Selling_df_match.at[index, '牌号匹配率'] = high_good_spec_match_rate
                Selling_df_match.at[index, '税收分类编码'] = high_standard_good_code
                temp_str = ''
                for item in possible_list:
                    temp_str += str(item)
                Selling_df_match.at[index, '可能的牌号品名'] = temp_str

                if high_good_name_match_rate < 50 or high_good_spec_match_rate < 50:
                    Selling_df_match.at[index, '潜在问题'] = '匹配率过低'
        except:
            pass

    # In[23]:


    try:
        warn_num = len(Selling_df_match.loc[Selling_df_match['潜在问题'] == '匹配率过低'])
        if warn_num != 0:
            print('警告有{}处匹配率过低!!'.format(warn_num))
    except:
        pass


    # # 发票模板

    # In[24]:


    Buyer_dict = dict(zip(Buyer_df['客户名称'], Buyer_df['统一社会信用代码/纳税人识别号']))


    # In[25]:


    pay_selling_df = Selling_df_match.loc[Selling_df_match['是否要开'] == '是']


    # In[26]:


    for rows in pay_selling_df.iterrows():
        index, columns = rows
        try:
            pay_selling_df.at[index, '统一社会信用代码/纳税人识别号'] = Buyer_dict[columns['客户']]
        except:
            pay_selling_df.at[index, '统一社会信用代码/纳税人识别号'] = 0


    # In[27]:


    for rows in Selling_df_match.iterrows():
        index, columns = rows
        try:
            Selling_df_match.at[index, '统一社会信用代码/纳税人识别号'] = Buyer_dict[columns['客户']]
        except:
            Selling_df_match.at[index, '统一社会信用代码/纳税人识别号'] = 0


    # In[28]:


    if list(pay_selling_df['统一社会信用代码/纳税人识别号']).count(0) > 0:
        print('警告：以下客户纳税人识别号不全')
        lst = pay_selling_df.loc[pay_selling_df['统一社会信用代码/纳税人识别号'] == 0]
        print(set(lst['客户']))


    # In[29]:


    Selling_df_match.to_excel('完整未开票文件(结果).xlsx')


    # In[30]:


    pay_selling_df.to_excel('只含需要开票文件(结果).xlsx')


# # 发票模板2

# In[ ]:
elif selection == 2:

    pay_selling_df = pd.read_excel('只含需要开票文件(结果).xlsx', dtype=str)


    # In[31]:


    result_basic = pd.DataFrame([['发票流水号', '发票类型', '特定业务类型', '是否含税', '受票方自然人标识', '购买方名称', '证件类型',
           '购买方纳税人识别号', '购买方地址', '购买方电话', '购买方开户银行', '购买方银行账号', '备注',
           '是否展示购买方银行账号', '销售方开户行', '销售方银行账号', '是否展示销售方银行账号', '购买方邮箱', '购买方经办人姓名',
           '购买方经办人证件类型', '购买方经办人证件号码', '经办人国籍(地区)', '经办人自然人纳税人识别号',
           '放弃享受减按1%征收率\n原因', '收款人', '复核人']])


    # In[32]:


    result_detail = pd.DataFrame([['发票流水号', '项目名称', '商品和服务税收编码', '规格型号', '单位', '数量', '单价', '金额', '税率',
           '折扣金额', '是否使用优惠政策', '优惠政策类型', '即征即退类型', 'index']])


    # In[33]:


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


    # In[34]:


    result_basic = result_basic.drop_duplicates(subset=0)


    # In[35]:


    result_basic.to_excel('1-发票基本信息.xlsx')


    # In[36]:


    for rows in pay_selling_df.iterrows():
        index, value = rows
        temp = pd.DataFrame([[value['销售订单号'], value['标准品名'], value['税收分类编码'], value['标准牌号'],
                              '吨', value['应开重量'], value['含税单价'], value['应开金额'], 0.13]])
        result_detail = pd.concat([result_detail, temp])


    # In[37]:


    result_detail.to_excel('2-发票明细信息.xlsx')

