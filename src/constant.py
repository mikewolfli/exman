#!/usr/bin/python3
#-*- coding:utf-8 -*-
'''
Created on 2016年7月21日

@author: mikewolfli
'''

tree_items=['数据同步','历史数据','走势图','数据分析']

#历史行情-表头对照
his_data_para_dic={
              'code': '股票代码', #即6位数字代码，或者指数代码（sh=上证指数 sz=深圳成指 hs300=沪深300指数 sz50=上证50 zxb=中小板 cyb=创业板）
              'start': '开始日期', #格式YYYY-MM-DD
              'end': '结束日期', #格式YYYY-MM-DD
              'ktype': '数据类型', #D=日k线 W=周 M=月 5=5分钟 15=15分钟 30=30分钟 60=60分钟，默认为D
              'retry_count':'当网络异常后重试次数', #默认为3
              'pause': '重试时停顿秒数',# 默认为0
              }

his_data_dic = {
                       'date': '日期',
                       'open':'开盘价',
                       'high':'最高价',
                       'close':'收盘价',
                       'low':'最低价',
                       'volume':'成交量',
                       'price_change':'价格变动',
                       'p_change':'涨跌幅',
                       'ma5':'5日均价',
                       'ma10':'10日均价',
                       'ma20':'20日均价',
                       'v_ma5':'5日均量',
                       'v_ma10':'10日均量',
                       'v_ma20':'20日均量',
                       'turnover':'换手率',#[注:指数无此项]
                       }

#复权数据
h_data_para_dic = {
                   'code':'股票代码',# string, e.g. 600848
                   'start':'开始日期',# string, format:YYYY-MM-DD 为空时取当前日期
                   'end':'结束日期', #string, format:YYYY-MM-DD 为空时取去年今日
                   'autype':'复权类型',#string,，qfq-前复权 hfq-后复权 None-不复权，默认为qfq
                   'index':'是否是大盘指数',#Boolean，，默认为False
                   'retry_count':'如遇网络等问题重复执行的次数',# int, 默认3,
                   'pause':'重复请求数据过程中暂停的秒数',#int, 默认 0,，防止请求间隔时间太短出现的问题
                   }

h_data_dic = {
              'date':'交易日期',# (index)
              'open':'开盘价',
              'high':'最高价',
              'close':'收盘价',
              'low':'最低价',
              'volume':'成交量',
              'amount':'成交金额',
              }

#实时行情 一次性获取当前交易所有股票的行情数据（如果是节假日，即为上一交易日，结果显示速度取决于网速）
real_price_all_dic = {
                      'code':'代码',
                      'name':'名称',
                      'changepercent':'涨跌幅',
                      'trade':'现价',
                      'open':'开盘价',
                      'high':'最高价',
                      'low':'最低价',
                      'settlement':'昨日收盘价',
                      'volume':'成交量',
                      'turnoverratio':'换手率',
              }

'''
历史分笔
获取个股以往交易历史的分笔数据明细，通过分析分笔数据，可以大致判断资金的进出情况。在使用过程中，对于获取股票某一阶段的历史分笔数据，需要通过参入交易日参数并append到一个DataFrame或者直接append到本地同一个文件里。历史分笔接口只能获取当前交易日之前的数据，当日分笔历史数据请调用get_today_ticks()接口或者在当日18点后通过本接口获取。
'''
tick_data_para_dic = {
                      'code':'股票代码',#即6位数字代码
                      'date':'日期',#格式YYYY-MM-DD
                      'retry_count':'重复次数',#int, 默认3,如遇网络等问题重复执行的次数
                      'pause':'暂停的秒数', #int, 默认 0,重复请求数据过程中暂停的秒数，防止请求间隔时间太短出现的问题
                      }

tick_data_dic = {
                 'time':'时间',
                 'price':'成交价格',
                 'change':'价格变动',
                 'volume':'成交手',
                 'amount':'成交金额(元)',
                 'type':'买卖类型',#【买盘、卖盘、中性盘】
                 
                 }

'''
实时分笔
获取实时分笔数据，可以实时取得股票当前报价和成交信息，其中一种场景是，写一个python定时程序来调用本接口（可两三秒执行一次，性能与行情软件基本一致），然后通过DataFrame的矩阵计算实现交易监控，可实时监测交易量和价格的变化。
'''
realtime_quotes_para_dic = { 'symbols':'股票代码'}#6位数字股票代码，或者指数代码（sh=上证指数 sz=深圳成指 hs300=沪深300指数 sz50=上证50 zxb=中小板 cyb=创业板） 可输入的类型：str、list、set或者pandas的Series对象


realtime_quotes_dic = {
                       }

#大盘指数实时行情列表-表头对照
index_dic={
             'code':'指数代码',
             'name':'指数名称',
             'change':'涨跌幅',
             'open':'开盘点位',
             'preclose':'昨日收盘点位',
             'close':'收盘点位',
             'high':'最高点位',
             'low':'最低点位',
             'volume':'成交量(手)',
             'amount':'成交金额（亿元）',
             }