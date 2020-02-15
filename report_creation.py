# -*- coding: utf-8 -*-

from openpyxl_simple_service import SimpleExcelService
from docx_simple_service import SimpleDocxService

e = SimpleExcelService()
e.open_wb('users.xlsx')
e.open_ws('Sheet1')
# users.xslxから値を取得
row_result = e.get_rows(6, 1)

# document作成
docx = SimpleDocxService()
date = input('Enter the date you want to create')
for row in row_result:
    # 名前
    name = row[0]
    # 機能訓練Ⅰ
    one_long_goal = row[2]
    one_long_level = row[3]
    one_short_goal = row[4]
    one_short_level = row[5]
    # 機能訓練Ⅱ
    two_long_goal = row[6]
    two_long_level = row[7]
    two_short_goal = row[8]
    two_short_level = row[9]
    # 評価
    evaluate = row[10]
    # 表題
    docx.add_head("機能訓練報告書", 0)
    # 御中とかそんなの
    docx.open_text()
    docx.add_text('石川内科胃腸科医院デイケア　切中様御中')
    docx.close_text()
    # 日付、発信元、住所
    docx.open_text()
    docx.add_text('{}\n'.format(date))
    docx.add_text('発信者 岡本荘デイサービスセンター 下大寺\n')
    docx.add_text('〒761-8047　高松市岡本町527-1\n')
    docx.add_text('TEL 087-885-2828\n')
    docx.paragraph_alignment_right()
    docx.close_text()
    # 礼文的なやつ
    docx.open_text()
    docx.add_text(
        'いつもお世話になっております。\n' \
        '各月における評価内容や、目標の達成度合いについて報告をさせていただきます。' \
        'ご確認のほどよろしくお願いします。また、当該利用者の目標、訓練内容の継続、' \
        'ADL及びIADLの改善状況を含めた目標の見直しや、訓練内容の変更など、お手数では' \
        'ありますが、電話等で助言のほどよろしくお願いいたします。'
    )
    docx.close_text()
    # 利用者氏名
    name = name
    docx.add_head('利用者氏名：{}'.format(name), 1)
    # 機能訓練Ⅰ
    docx.add_head('個別機能訓練Ⅰ', 1)
    docx.open_text()
    docx.add_text('【長期目標】：{}\n'.format(one_long_goal))
    docx.add_text('　【達成度】：{}\n'.format(one_long_level))
    docx.add_text('【短期目標】：{}\n'.format(one_short_goal))
    docx.add_text('　【達成度】：{}'.format(one_short_level))
    docx.close_text()
    # 機能訓練Ⅱ
    docx.add_head('個別機能訓練Ⅱ', 1)
    docx.open_text()
    docx.add_text('【長期目標】：{}\n'.format(two_long_goal))
    docx.add_text('　【達成度】：{}\n'.format(two_long_level))
    docx.add_text('【短期目標】：{}\n'.format(two_short_goal))
    docx.add_text('　【達成度】：{}'.format(two_short_level))
    docx.close_text()
    # 評価内容
    docx.add_head('評価内容', 1)
    docx.open_text()
    docx.add_text('{}'.format(evaluate))
    # 改行
    docx.add_page_break()

# 保存して終了
docx.save("demo.docx")
print("complete")
