import config
import os
import sys
from function import Comments, OrderRecord, OrderList


def translate_params():
    goods_type = '区分柄, 柄类型: %s' % '、'.join(config.goods_type) if config.goods_type else '不区分柄'
    is_template = '是' if os.path.exists(os.path.join(config.path_project, '模板.xlsx')) else '否'
    return '文件保存路径： %s; 模板文件是否存在: %s, 最小起排数量: %d, 最大数量: %d, %s' % (
        os.path.abspath(config.path_project), is_template, config.quantity_min, config.quantity_max, goods_type)


def main():
    if not os.path.exists(os.path.join(config.path_project, '模板.xlsx')):
        print('模板文件不存在, 请检查配置, 目标路径: %s' % os.path.abspath(os.path.join(config.path_project, '模板.xlsx')))
        sys.exit()
    # 确认输入参数的含义
    print('请确认参数是否正确')
    print(translate_params())
    step = input('选择执行步骤:\n1-全部执行, 2-格式化评论并生成排表, 3-生成排表\n')
    if step == '1':
        # 执行获取评论的脚本
        print('开始抓取相册评论')
        comments = Comments(config.album, config.login, config.path_project)
        dataframe = comments.get_comments_concat()
        comments.export_xlsx(dataframe)
        print('共计%d条评论, 新增评论%d条' % (len(dataframe), comments.count_new))
    if step in ('1', '2'):
        input('请检查评论是否符合规范, 文件路径: %s, 列名: 相册评论（可修改）, 按任意键确认' %
                     os.path.abspath(os.path.join(config.path_project, '相册评论.xlsx')))
        # 执行规范化评论的脚本
        print('开始生成规范化评论')
        order = OrderRecord(config.myname, config.path_project, config.path_goods, config.goods_type)
        dataframe = order.get_comments_normalize()
        dataframe_error = dataframe[dataframe['评论规范化'] == 'Error']
        order.export_xlsx(dataframe)
        print('已对%d条新增评论进行规范化, 共计%d条报错, 错误记录如下:' % (order.count_new, len(dataframe_error)))
        print(dataframe_error)
    # 输出排表
    print('开始生成排表')
    orderlist = OrderList(config.path_goods, config.path_project, config.quantity_min, config.quantity_max,
                          config.goods_type)
    orderlist.export_xlsx()
    print('完成')


if __name__ == '__main__':
    main()
