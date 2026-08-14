# -*- coding: utf-8 -*-
"""
申万/中信一级行业板块分类排序模块
供各绘图脚本导入使用
"""

# 板块分类定义
SW_SECTOR_MAP = {
    '上游资源': ['石油石化', '煤炭', '有色金属', '基础化工', '钢铁', '农林牧渔'],
    '中游制造': ['电力设备', '国防军工', '机械设备', '建筑装饰', '建筑材料',
                 '交通运输', '轻工制造', '公用事业', '环保'],
    '下游消费': ['汽车', '商贸零售', '家用电器', '食品饮料', '纺织服饰',
                '医药生物', '社会服务', '美容护理'],
    'TMT': ['电子', '计算机', '通信', '传媒'],
    '金融地产': ['银行', '非银金融', '房地产'],
    '综合': ['综合'],
}

ZX_SECTOR_MAP = {
    '上游资源': ['石油石化', '煤炭', '有色金属', '基础化工', '钢铁', '农林牧渔'],
    '中游制造': ['电力设备及新能源', '国防军工', '机械', '建筑', '建材',
                 '交通运输', '轻工制造', '电力及公用事业'],
    '下游消费': ['汽车', '商贸零售', '家电', '食品饮料', '纺织服装',
                '医药', '消费者服务'],
    'TMT': ['电子', '通信', '计算机', '传媒'],
    '金融地产': ['银行', '非银行金融', '房地产', '综合金融'],
    '综合': ['综合'],
}


def reindex_by_sector(df, classification='sw'):
    """
    将DataFrame的行按板块分类顺序重排
    classification: 'sw' 申万, 'zx' 中信
    """
    sector_map = SW_SECTOR_MAP if classification == 'sw' else ZX_SECTOR_MAP
    ordered_inds = []
    for sector, inds in sector_map.items():
        ordered_inds.extend([i for i in inds if i in df.index])
    unmatched = [i for i in df.index if i not in ordered_inds]
    return df.reindex(ordered_inds + unmatched)
