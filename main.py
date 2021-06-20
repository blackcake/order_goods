import config
import os
import openpyxl
import sys
from function import Comments, OrderRecord, OrderList


def translate_params():
    goods_type = '区分柄, 柄类型: %s' % '、'.join(config.goods_type) if config.goods_type else '不区分柄'
    is_template = '是' if os.path.exists(os.path.join(config.project, '模板.xlsx')) else '否'
    return '文件保存路径： %s; 模板文件是否存在: %s, 最小起排数量: %d, 最大数量: %d, %s' % (
        os.path.abspath(config.project), is_template, config.quantity_min, config.quantity_max, goods_type)


def check():
    if not os.path.exists(os.path.join(config.project, '模板.xlsx')):
        print('模板文件不存在, 请检查配置, 目标路径: %s' % os.path.abspath(os.path.join(config.project, '模板.xlsx')))
        sys.exit()
    workbook = openpyxl.load_workbook(os.path.join(config.project, '模板.xlsx'))
    worksheet = workbook.active
    cols = ('角色', '柄类型', '单价', '余量') if config.goods_type else ('角色', '单价', '余量')
    cols_sheet = next(worksheet.values)
    for col in cols:
        if col not in cols_sheet:
            print('模板文件错误，不包含 %s 列' % col)
            sys.exit()


def main():
    # 检查模板格式
    check()
    # 确认输入参数的含义
    print('请确认参数是否正确')
    print(translate_params())
    step = input('选择执行步骤:\n1-全部执行, 2-格式化评论并生成排表, 3-生成排表\n')
    if step == '1':
        # 执行获取评论的脚本
        print('开始抓取相册评论')
        comments = Comments(config.album, config.login, config.project)
        comments.export()
        print('相册评论共计%d条, 原有评论%d条, 新增评论%d条' % (comments.count_comment, comments.count_history, comments.count_new))
    if step in ('1', '2'):
        input('请检查评论是否符合规范, 文件路径: %s, 列名: 相册评论（可修改）, 按回车确认' %
              os.path.abspath(os.path.join(config.project, '相册评论.xlsx')))
        # 执行规范化评论的脚本
        print('开始生成规范化评论')
        order = OrderRecord(config.myname, config.project, config.path_goods, config.goods_type)
        order.export()
        print('评论共计%d条, 已对%d条新增评论进行规范化, 共计%d条报错' % (order.count_comment, order.count_new, len(order.dataframe_error)))
        if not order.dataframe_error.empty:
            print(', 错误记录如下:')
            print(order.dataframe_error)
    # 输出排表
    print('开始生成排表')
    orderlist = OrderList(config.path_goods, config.project, config.quantity_min, config.quantity_max,
                          config.goods_type)
    orderlist.export()
    print('完成')


if __name__ == '__main__':
    main()
