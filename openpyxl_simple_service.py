# -*- coding: utf-8 -*-
"""
Created on Sat Feb 15 14:46:04 2020

@author: Okamoto-DRIHA001
"""

import openpyxl as opxl
import numpy as np


class SimpleExcelService:

    def open_wb(self, file_path):
        # ワークブックを開く
        self.wb = opxl.load_workbook(file_path)

    def open_ws(self, sheet_name):
        # ワークシートを開く
        self.ws = self.wb[sheet_name]

    def get_max_row(self):
        # 入力されている最大行を返す
        return self.ws.max_row

    def get_max_column(self):
        # 入力されている最大列を返す
        return self.ws.max_column

    def get_rows(self, min_row, min_col):
        # 指定したセルから行単位で値を取得する
        row_result = [row for row in self.ws.iter_rows(min_row=min_row, min_col=min_col, values_only=True)]
        return row_result

    def save_wb(self, save_path):
        # ワークブックを保存する
        self.wb.save(save_path)
        print("{}: Saved Successfully".format(save_path))

    def close_wb(self):
        # ワークブックを閉じる
        self.wb.close()
        print("closed workbook")
