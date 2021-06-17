# -*- coding: utf-8 -*-
import json
import openpyxl
import os
import pandas
import re
import requests
import time
from openpyxl.utils.dataframe import dataframe_to_rows


class Comments:
    def __init__(self, album, login, path_project):
        self._album = album
        self._login = login
        self._path_project = path_project
        self._path_comments = os.path.join(path_project, '相册评论.xlsx')
        self.count_new = 0

    def get_comments_history(self):
        if not os.path.exists(self._path_comments):
            return pandas.DataFrame([])
        return self._read_xlsx()

    def _read_xlsx(self):
        workbook = openpyxl.load_workbook(self._path_comments)
        worksheet = workbook.active
        data = worksheet.values
        cols = next(data)
        dataframe = pandas.DataFrame(data, columns=cols)
        return dataframe

    def _get_data(self, start):
        url = 'https://h5.qzone.qq.com/proxy/domain/u.photo.qzone.qq.com/cgi-bin/upp/qun_list_photocmt_v2'
        query_string = {
            'uin': self._album['uin'],
            'hostUin': self._album['hostUin'],
            'start': '{:d}'.format(start),
            'num': '10',
            'order': '0',
            'topicId': self._album['topicId'],
            'format': 'jsonp',
            'inCharset': 'utf-8',
            'outCharset': 'utf-8',
            'ref': 'qunphoto',
            'random': self._login['random'],
            'g_tk': self._login['g_tk'],
            'qzonetoken': self._login['qzonetoken']
        }
        header = {
            'cookie': self._login['cookie']
        }
        data = requests.request("GET", url, headers=header, params=query_string).text
        return json.loads(data[10: -2])['data']

    def get_comments_now(self):
        data = self._get_data(0)
        comments = data['comments']
        for i in range(10, data['total'], 10):
            time.sleep(1)
            comments.extend(self._get_data(i)['comments'])
        dict_comments = {
            'content': [c['content'] for c in comments],
            'id': [c['id'] for c in comments],
            'postTime': [c['postTime'] for c in comments],
            'name': [c['poster']['name'] for c in comments],
        }
        dataframe = pandas.DataFrame(dict_comments)[['name', 'id', 'postTime', 'content']]
        dataframe.insert(len(dataframe.columns), 'content_correct', dataframe['content'])
        return dataframe

    def get_comments_concat(self):
        df_history = self.get_comments_history()
        df_now = self.get_comments_now()
        if df_history.empty:
            return df_now
        row = df_history.loc[df_history.index[-1]].tolist()
        i = -1
        for i in df_now.index:
            if df_now.loc[i].tolist() == row:
                break
        self.count_new = len(df_now) - i - 1
        return pandas.concat([df_history, df_now[i+1:]])

    def export_xlsx(self, dataframe):
        dataframe = dataframe.rename(columns={'name': '群昵称', 'postTime': '评论时间', 'content': '相册评论（请勿修改）',
                                              'content_correct': '相册评论（可修改）'})
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        for row in dataframe_to_rows(dataframe, index=False, header=True):
            sheet.append(row)
        if not os.path.exists(self._path_project):
            os.mkdir(self._path_project)
        workbook.save(self._path_comments)


class OrderRecord:
    def __init__(self, myname, path_project, path_goods, goods_type):
        self._myname = myname
        self._path_comments = os.path.join(path_project, '相册评论.xlsx')
        self._df_goods = self._read_xlsx(path_goods)
        self._goods_type = goods_type
        self.count_new = 0

    def get_comments(self):
        if not os.path.exists(self._path_comments):
            raise ValueError('找不到文件')
        dataframe = self._read_xlsx()
        if '圈名' not in dataframe.columns:
            dataframe.insert(len(dataframe.columns), '圈名', None)
        if '评论规范化' not in dataframe.columns:
            dataframe.insert(len(dataframe.columns), '评论规范化', None)
        return dataframe

    def _read_xlsx(self, path=None):
        if not path:
            path = self._path_comments
        workbook = openpyxl.load_workbook(path)
        worksheet = workbook.active
        data = worksheet.values
        cols = next(data)
        dataframe = pandas.DataFrame(data, columns=cols)
        return dataframe

    def _normalize_user(self, string):
        if string == self._myname['name']:
            return self._myname['nickname']
        # 规范化用户名称
        return re.split('[- （(@【—_–，。]', string)[0]

    def _match_goods(self, string, dataframe_goods, number):
        list_role_norm = self._match_goods_type(string)
        if list_role_norm == 'Error':
            return 'Error'
        result = []
        for string, role_type in list_role_norm:
            list_goods = []
            for i in dataframe_goods.index:
                if dataframe_goods.loc[i, '别名'] in string and dataframe_goods.loc[i, '全称'] not in list_goods:
                    list_goods.append(','.join([dataframe_goods.loc[i, '全称'], role_type, number]))
                    string = string.replace(dataframe_goods.loc[i, '别名'], '')
                if not string:
                    break
            if list_goods:
                result.append(';\n'.join(list_goods))
        return ';\n'.join(result)

    def _normalize_goods(self, string, dataframe):
        # 规范化用户点单要求的字符串，如vk和kn各1，改成valkyrie1+knights1
        # 先不考虑队走的和转单的
        # 暂时只考虑数量后置的
        if self._skip_rows(string):
            return 'Error'
        # all有3个字符，替换；兔团和仁兔的缩写容易重，替换
        string = string.strip().lower().replace('all', '凹').replace('兔团', 'rabbit').replace('.', '').replace('。', '')
        # 如果角色和数字对不起来或者识别到的角色数量为0，返回错误，手动填
        list_goods = [row for row in re.split('[0-9凹|余]+', string) if row]
        list_number = [row for row in re.split('[^0-9凹|余]+', string) if row]
        if len(list_goods) < len(list_number) or not list_goods:
            return 'Error'
        # 把凹余/余凹/余/凹统一成一种表述
        list_number = ['凹' if '余' in row else row for row in list_number]
        result_match = [self._match_goods(list_goods[i], dataframe, list_number[i]) for i in range(len(list_number))]
        order = ';\n'.join([r for r in result_match if r])
        if 'Error' in order:
            return 'Error'
        return order

    def _skip_rows(self, string):
        if not string:
            return 1
        if '转' in string or '接' in string or '撤' in string:
            return 1
        return 0

    def _match_goods_type(self, string):
        if not self._goods_type:
            return [[string]]
        list_role = [row for row in re.split('[%s]+' % ''.join(set(list(''.join(self._goods_type)))), string.strip()) if row]
        list_type = [row for row in re.split('[^%s]+' % ''.join(set(list(''.join(self._goods_type)))), string.strip()) if row]
        if len(list_role) < len(list_type):
            return 'Error'
        list_role_norm = []
        for i in range(len(list_role)):
            if i >= len(list_type):
                list_role_norm.append([list_role[i], 'all'])
                continue
            role_type = [t for t in self._goods_type if t in list_type[i]]
            if len(role_type) == len(self._goods_type):
                list_role_norm.append([list_role[i], 'all'])
            else:
                list_role_norm.extend([[list_role[i], t] for t in role_type])
        return list_role_norm

    def get_comments_normalize(self):
        dataframe = self.get_comments()
        dataframe_empty = dataframe[dataframe['评论规范化'].isnull()].copy()
        self.count_new = len(dataframe_empty)
        dataframe_empty['圈名'] = dataframe_empty['群昵称'].apply(self._normalize_user)
        dataframe_empty['评论规范化'] = dataframe_empty['相册评论（可修改）'].apply(self._normalize_goods,
                                                                      args=(self._df_goods,))
        for i in dataframe_empty.index:
            dataframe.loc[i, '圈名'] = dataframe_empty.loc[i, '圈名']
            dataframe.loc[i, '评论规范化'] = dataframe_empty.loc[i, '评论规范化']
        return dataframe

    def export_xlsx(self, dataframe):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        for row in dataframe_to_rows(dataframe, index=False, header=True):
            sheet.append(row)
        workbook.save(self._path_comments)


class OrderList:
    def __init__(self, path_goods, path_project, quantity_min, quantity_max, goods_type):
        self._path_project = path_project
        self._path_comments = os.path.join(path_project, '相册评论.xlsx')
        self._df_team = self._read_xlsx(path_goods, '队伍')
        self._path_template = os.path.join(path_project, '模板.xlsx')
        self._quantity_min = quantity_min
        self._quantity_max = quantity_max
        self._goods_type = goods_type

    def get_comments(self):
        if not os.path.exists(self._path_comments):
            raise ValueError('找不到文件')
        dataframe = self._read_xlsx(self._path_comments, 'Sheet')[['圈名', 'id', '评论规范化']].rename(
            columns={'圈名': 'name', '评论规范化': 'order'})
        return dataframe[dataframe['order'] != 'Error'].reset_index(drop=True)

    def _read_xlsx(self, path, sheet=None):
        workbook = openpyxl.load_workbook(path)
        if sheet:
            worksheet = workbook[sheet]
        else:
            worksheet = workbook.active
        data = worksheet.values
        cols = next(data)
        dataframe = pandas.DataFrame(data, columns=cols)
        return dataframe

    def get_comments_detail(self):
        dataframe = self.get_comments()
        line_split = dataframe['order'].str.split(';\n', expand=True).stack().reset_index(level=1, drop=True)
        dataframe = dataframe[['name', 'id']].join(line_split.rename('order'))
        if self._goods_type:
            dataframe[['role', 'type', 'quantity']] = dataframe['order'].str.split(',', 2, expand=True)
        else:
            dataframe[['role', 'quantity']] = dataframe['order'].str.split(',', 1, expand=True)
        dataframe['is_all'] = dataframe['quantity'].apply(lambda x: 1 if x == '凹' else 0)
        dataframe['quantity'] = dataframe['quantity'].apply(lambda x: int(x) if x != '凹' else self._quantity_max)
        return dataframe.reset_index(drop=False)

    def _match_team(self, dataframe):
        dataframe_match = pandas.merge(dataframe, self._df_team, left_on='role', right_on='队伍', how='left')
        for i in dataframe_match.index:
            if not pandas.isnull(dataframe_match.loc[i, '角色']):
                dataframe_match.loc[i, 'role'] = dataframe_match.loc[i, '角色']
        dataframe_match = dataframe_match.drop(['队伍', '角色'], axis=1)
        return dataframe_match

    def _match_type(self, dataframe):
        if not self._goods_type:
            return dataframe
        df_template = self._read_xlsx(self._path_template)
        dataframe_goods_all = pandas.merge(dataframe[dataframe['type'] == 'all'], df_template, left_on='role',
                                           right_on='角色', how='inner')[['index', 'name', 'role', '柄类型', 'quantity',
                                                                        'is_all']].rename(columns={'柄类型': 'type'})
        dataframe_goods_single = dataframe[dataframe['type'] != 'all'][['index', 'name', 'role', 'type', 'quantity',
                                                                        'is_all']]
        dataframe = pandas.concat([dataframe_goods_all, dataframe_goods_single])
        return dataframe.sort_values('index').drop('index', axis=1).reset_index(drop=True)

    def _calculate_real_number(self, dataframe):
        if self._goods_type:
            column_groupby = ['role', 'type']
        else:
            column_groupby = ['role']
        dataframe = dataframe.copy()
        dataframe.insert(0, 'namelist', dataframe.apply(lambda x: ','.join(
            [x['name']] * int(x['quantity']/self._quantity_min)), axis=1))
        dataframe = dataframe.groupby(column_groupby).agg({'quantity': 'sum', 'namelist': lambda x: ','.join(x)}
                                                          ).reset_index(drop=False)
        for i in dataframe.index:
            if dataframe.loc[i, 'quantity'] >= self._quantity_max:
                dataframe.loc[i, 'quantity'] = self._quantity_max
                dataframe.loc[i, 'namelist'] = ','.join(
                    dataframe.loc[i, 'namelist'].split(',')[:int(self._quantity_max/self._quantity_min)])
        return dataframe

    def _match_price(self, dataframe):
        df_template = self._read_xlsx(self._path_template)
        if self._goods_type:
            column_left, column_right, column_rename = ['role', 'type'], ['角色', '柄类型'], {}
        else:
            column_left, column_right = ['role'], ['角色']
        dataframe = dataframe.groupby(['name']+column_left).agg({'quantity': 'sum'}).reset_index(drop=False)
        dataframe = pandas.merge(dataframe, df_template, left_on=column_left, right_on=column_right, how='inner')
        dataframe.insert(0, '总价', dataframe['单价'] * dataframe['quantity'])
        return dataframe[['name']+column_right+['quantity', '单价', '总价']].rename(columns={'name': '圈名',
                                                                                         'quantity': '数量'})

    def get_comments_calc(self):
        dataframe = self.get_comments_detail()
        dataframe = self._match_team(dataframe)
        dataframe = self._match_type(dataframe)
        dataframe_order = self._calculate_real_number(dataframe)
        dataframe_price = self._match_price(dataframe)
        return dataframe_order, dataframe_price

    def _match_index(self, dataframe):
        df_template = self._read_xlsx(self._path_template).reset_index(drop=False)
        if self._goods_type:
            column_left, column_right, column_rename = ['role', 'type'], ['角色', '柄类型'], {}
        else:
            column_left, column_right = ['role'], ['角色']
        dataframe = pandas.merge(dataframe, df_template, left_on=column_left, right_on=column_right, how='inner')
        dataframe.index = dataframe['index'] + 2
        return dataframe

    def export_xlsx(self):
        dataframe_order, dataframe_price = self.get_comments_calc()
        dataframe_order = self._match_index(dataframe_order)
        workbook = openpyxl.load_workbook(self._path_template)
        sheet_role = workbook.active
        columns = next(sheet_role.values)
        index_role, index_sum, index_number = columns.index('角色')+1, columns.index('余量')+1, columns.index('余量')+2
        for i in range(2, sheet_role.max_row+1):
            if i in dataframe_order.index:
                sheet_role.cell(row=i, column=index_sum).value = self._quantity_max - dataframe_order.loc[i, 'quantity']
                namelist = dataframe_order.loc[i, 'namelist'].split(',')
                for j in range(len(namelist)):
                    sheet_role.cell(row=i, column=j+index_number).value = namelist[j]
            elif sheet_role.cell(row=i, column=index_role).value:
                sheet_role.cell(row=i, column=index_sum).value = self._quantity_max
        sheet_person = workbook.create_sheet('个人')
        for row in dataframe_to_rows(dataframe_price, index=False, header=True):
            sheet_person.append(row)
        workbook.save(os.path.join(self._path_project, '排表.xlsx'))
