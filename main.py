import config
import os
import sys
from function import Comments, OrderRecord, OrderList, read_xlsx


def print_params(project, goods_type, quantity_min, quantity_max):
    goods_type = '区分柄, 柄类型: %s' % '、'.join(goods_type) if goods_type else '不区分柄'
    print('文件保存路径： %s; 最小起排数量: %d, 最大数量: %d, %s' % (
        os.path.abspath(project), quantity_min, quantity_max, goods_type))


def check_params(project):
    # 角色清单
    if not os.path.exists(config.path_goods):
        print('角色清单.xlsx文件不存在, 请检查配置, 目标路径: %s' % os.path.abspath(config.path_goods))
        sys.exit()
    # 模板
    path_template = os.path.join(project, '模板.xlsx')
    if not os.path.exists(path_template):
        print('模板文件不存在, 请检查配置, 目标路径: %s' % os.path.abspath(path_template))
        sys.exit()
    try:
        df_template = read_xlsx(path_template, '排表')
    except:
        print('读取模板文件的"排表"sheet失败')
        sys.exit()
    for col in ('角色', '柄类型', '单价', '余量'):
        if col not in df_template.columns:
            print('模板文件错误，"排表"sheet不包含"%s"列' % col)
            sys.exit()
    try:
        df_params = read_xlsx(path_template, '参数')
    except:
        print('读取模板文件的"参数"sheet失败')
        sys.exit()
    for col in ('最小起排数量', '最大配比数', 'topicId'):
        if col not in df_params['参数名'].to_list():
            print('模板文件错误，"参数"sheet不包含"%s"列' % col)
            sys.exit()


def read_params(project):
    path_template = os.path.join(project, '模板.xlsx')
    # 柄类型直接读模板
    df_template = read_xlsx(path_template, '排表')
    goods_type = [t for t in set(df_template['柄类型'].to_list()) if t]
    # 其他参数
    df = read_xlsx(path_template, '参数')
    params = {df.loc[i, '参数名']: df.loc[i, '参数值'] for i in df.index}
    return goods_type, int(params['最小起排数量']), int(params['最大配比数']), params['topicId']


def main():
    print('请先确认QQ相册的headers参数已修改完成, 配置文件路径: %s' % os.path.abspath('./config.py'))
    project = input('请输入项目名称: ')
    # 检查和参数格式
    check_params(project)
    goods_type, quantity_min, quantity_max, topicId = read_params(project)
    # 确认输入参数的含义
    print('请确认参数是否正确')
    print_params(project, goods_type, quantity_min, quantity_max)
    step = input('选择执行步骤:\n1-全部执行, 2-格式化评论并生成排表\n')
    if step == '1':
        # 执行获取评论的脚本
        print('开始抓取相册评论')
        comments = Comments(config.album, config.login, project, topicId)
        comments.export()
        print('相册评论抓取成功, 文件路径: %s, 评论共计%d条, 原有评论%d条, 新增评论%d条' % (
            comments.path_comments, comments.count_comment, comments.count_history, comments.count_new))
    if step in ('1', '2'):
        # 执行规范化评论的脚本
        print('开始生成规范化评论')
        order = OrderRecord(config.myname, project, config.path_goods, goods_type, quantity_max)
        order.export()
        print('已对%d条新增评论进行规范化, 共计%d条报错' % (order.count_comment, len(order.dataframe_error)))
        if not order.dataframe_error.empty:
            print('错误记录如下:')
            print(order.dataframe_error[['圈名', '相册评论（可修改）']])
    # 输出排表
    print('开始生成排表')
    orderlist = OrderList(config.path_goods, project, quantity_min, quantity_max, goods_type)
    orderlist.export()
    print('完成')


if __name__ == '__main__':
    main()
