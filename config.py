# -*- coding: utf-8 -*-
# 只有第一次执行的时候需要修改
path_goods = r'F:\scripts\order_goods\角色清单.xlsx'
myname = {'name': '艾姆索赛德', 'nickname': '洛书'}

# 不同类型的谷子需要修改
path_project = r'.\拍立得v4'
quantity_min = 1
quantity_max = 10
album = {
    'uin': '584735693',
    'hostUin': '861783543',
    'topicId': '766122119|V51ZHlJI03Tcod0igfPv3DJGfw3O1ZS6|V5bCQA3NjYxMjIxMTluY7BgQ1aNFg!!'
}

# 每次执行都需要修改（如果是短时间内跑同一个群内的好几个相册，可以不用改）
login = {
    'qzonetoken': 'fd8eaa526a9c80176234c9b53d4e0e26f4c69e135d635823406e24961ebe11a2a3f52303698b6dd862b14729739f78d76ee8',
    'random': '0.0172708932308423',
    'g_tk': '1381611710',
    'cookie': 'pgv_pvid=1900968440; RK=1DbEKqqXEz; ptcz=3673bf3b210cc0fa15c5bb07d44020d51ffe3ab87d76aedc5b4abf32b86a05cd; tvfe_boss_uuid=c9dcc2104191e416; mobileUV=1_175e5a66015_16288; pgv_pvi=7925114880; eas_sid=K196S1S3R5I6Q2K2H0z0u7u895; ied_qq=o0584735693; o_cookie=584735693; __Q_w_s_hat_seed=1; pac_uid=1_584735693; uin=o0584735693; skey=@KGPKeZyzs; p_uin=o0584735693; pt4_token=61mTFo8nWuLMNfVydd2ciCS8pd9v2D2RQc-vbOrTFoI_; p_skey=ZZDgw5KgmBx*NpgWxLv7B-gZiROzhGBPw7orAhdblKU_; pgv_info=ssid=s5857669533'
}

# 如果有不同柄，需要修改
goods_type = ['花前', '花后']


'''
现有问题：
(放弃, 不想做)1. 不支持用中文数字 - 个位数还好，十二/三十三这种数字太难搞了，准备放弃
(放弃，手动调评论)2. 只支持时间顺序，不支持识别管理员优先级（可以手动调整相册评论.xlsx）
(懒得做了, 影响不大)3. 在有起排最小值的情况下，比如10起排，先5再5两条评论无法合并（可以手动调整相册评论.xlsx）
(完成)4. 跳过历史记录
(放弃, 规范数据要求)5. 无法识别数字前置的情况，目前正则判断的标准：角色（可以有一长串很多人的缩写）+数量+角色（再一长串）+数量，如零凛10薰晃20eve兔团30（可以手动调整相册评论规范化.xlsx）
(懒得做了，延后)6. 无法判断是否队走/对走
(不用改，添加到说明文档)7. 如果出现不认识的缩写，需要手动添加到goods.xlsx
(完成)8. 不支持同人不同柄
(完成)9. 排表最好可以根据模板填写结果
(懒得做了, 无限延后)10. 待添加：生成余量图，计算成多少对
(懒得做了)11. all用户单元格上色
'''