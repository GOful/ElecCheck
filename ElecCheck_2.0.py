import flet as ft
import threading
import openpyxl
import pandas as pd
import re
import math

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys

EDGE_DRIVER_PATH = r""


def main(page: ft.Page):
    page.title = "한전 원격검침 비교 (Flet)"
    page.bgcolor = "#f3f7f6"
    page.padding = 16
    page.horizontal_alignment = "stretch"
    page.vertical_alignment = "stretch"

    PRIMARY = "#4db6ac"
    PRIMARY_DARK = "#00867d"
    WARNING = "#d32f2f"

    state = {
        "workbook": None,
        "sheet_name": None,
        "initial_values": [],
    }

    log_view = ft.ListView(expand=1, spacing=4, auto_scroll=True)

    def log(msg: str):
        log_view.controls.append(ft.Text(msg, size=12))
        page.update()

    progress = ft.ProgressBar(width=400, value=0, bgcolor="#dce4e3")

    table = ft.DataTable(
        columns=[
            ft.DataColumn(ft.Text("변전소")),
            ft.DataColumn(ft.Text("9")),
            ft.DataColumn(ft.Text("10")),
            ft.DataColumn(ft.Text("11")),
            ft.DataColumn(ft.Text("12")),
            ft.DataColumn(ft.Text("13")),
            ft.DataColumn(ft.Text("14")),
            ft.DataColumn(ft.Text("15")),
        ],
        rows=[],
    )

    run_button = ft.ElevatedButton(
        "실행",
        disabled=True,
        bgcolor=PRIMARY,
        color="white",
        style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=8)),
    )

    file_picker = ft.FilePicker()
    page.overlay.append(file_picker)

    # 시트 모달
    sheet_radio = ft.RadioGroup(content=ft.Column([]))
    sheet_dialog = ft.AlertDialog(
        modal=True,
        title=ft.Text("시트 선택"),
        content=sheet_radio,
        actions=[],
        actions_alignment="end",
    )

    def close_sheet_dialog(e=None):
        sheet_dialog.open = False
        page.update()

    def load_sheet_to_table(sheet):
        cell_ranges = [
            ['B4', 'C6', 'C7', 'C8', 'C9', 'C10', 'C11', 'C12'],
            ['F4', 'G6', 'G7', 'G8', 'G9', 'G10', 'G11', 'G12'],
            ['J4', 'K6', 'K7', 'K8', 'K9', 'K10', 'K11', 'K12'],

            ['B17', 'C19', 'C20', 'C21', 'C22', 'C23', 'C24', 'C25'],
            ['F17', 'G19', 'G20', 'G21', 'G22', 'G23', 'G24', 'G25'],
            ['J17', 'K19', 'K20', 'K21', 'K22', 'K23', 'K24', 'K25'],

            ['B30', 'C32', 'C33', 'C34', 'C35', 'C36', 'C37', 'C38'],

            ['B44', 'C46', 'C47', 'C48', 'C49', 'C50', 'C51', 'C52'],
            ['F44', 'G46', 'G47', 'G48', 'G49', 'G50', 'G51', 'G52'],
            ['J44', 'K46', 'K47', 'K48', 'K49', 'K50', 'K51', 'K52'],

            ['B57', 'C59', 'C60', 'C61', 'C62', 'C63', 'C64', 'C65'],
            ['F57', 'G59', 'G60', 'G61', 'G62', 'G63', 'G64', 'G65'],
            ['J57', 'K59', 'K60', 'K61', 'K62', 'K63', 'K64', 'K65'],

            ['B70', 'C72', 'C73', 'C74', 'C75', 'C76', 'C77', 'C78'],
            ['F70', 'G72', 'G73', 'G74', 'G75', 'G76', 'G77', 'G78'],
            ['J70', 'K72', 'K73', 'K74', 'K75', 'K76', 'K77', 'K78'],

            ['B83', 'C85', 'C86', 'C87', 'C88', 'C89', 'C90', 'C91'],
        ]
        cell_values = [
            '설화명곡', '월배기지', '서부정류장', '반월당', '신천', '방촌', '안심',
            '문양기지', '대실', '성서산단', '죽전', '반고개', '대구은행', '만촌', '대공원', '사월', '영남대'
        ]

        rows = []
        state["initial_values"].clear()

        for cell_range, name in zip(cell_ranges, cell_values):
            row_vals = [name]
            for idx, cell_addr in enumerate(cell_range):
                if idx == 0:
                    continue
                row_vals.append(sheet[cell_addr].value)
            rows.append(
                ft.DataRow(
                    cells=[ft.DataCell(ft.Text(str(v) if v is not None else "")) for v in row_vals]
                )
            )
            state["initial_values"].append(row_vals)

        table.rows = rows
        page.update()

    def confirm_sheet_dialog(e):
        sn = sheet_radio.value
        if not sn:
            return
        sheet = state["workbook"][sn]
        state["sheet_name"] = sn
        load_sheet_to_table(sheet)
        run_button.disabled = False
        close_sheet_dialog()
        log(f"시트 '{sn}' 로드 완료")

    def open_sheet_dialog(sheetnames):
        sheet_radio.content.controls.clear()
        for sn in sheetnames:
            sheet_radio.content.controls.append(ft.Radio(value=sn, label=sn))
        if sheetnames:
            sheet_radio.value = sheetnames[0]

        sheet_dialog.actions = [
            ft.TextButton("취소", on_click=close_sheet_dialog),
            ft.ElevatedButton("확인", bgcolor=PRIMARY, color="white", on_click=confirm_sheet_dialog),
        ]

        page.dialog = sheet_dialog
        sheet_dialog.open = True
        page.update()

    # 파일 선택 끝났을 때
    def on_file_result(e: ft.FilePickerResultEvent):
        if not e.files:
            return
        file_path = e.files[0].path
        log(f"엑셀 불러오는 중: {file_path}")
        wb = openpyxl.load_workbook(file_path, read_only=True)
        state["workbook"] = wb

        # 최신 Flet에서는 이렇게 바로 띄워도 됨
        open_sheet_dialog(wb.sheetnames)

    file_picker.on_result = on_file_result

    def open_file_click(e):
        file_picker.pick_files(allow_multiple=False)

    # 웹 작업은 네가 쓰던 거 다시 붙이면 됨
    # (여기서는 생략할게, 위에서 이미 길게 있었으니까)

    top_bar = ft.Container(
        content=ft.Row(
            controls=[
                ft.ElevatedButton(
                    "엑셀 열기",
                    on_click=open_file_click,
                    bgcolor="white",
                    color=PRIMARY_DARK,
                    style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=8)),
                ),
                run_button,
            ],
            alignment=ft.MainAxisAlignment.START,
            spacing=12,
        ),
        padding=12,
        bgcolor=PRIMARY,
        border_radius=12,
    )

    table_wrap = ft.Container(
        content=table,
        bgcolor="white",
        border_radius=12,
        padding=8,
        expand=True,
    )

    logs_wrap = ft.Container(
        content=ft.Column([log_view, progress], expand=True),
        bgcolor="white",
        border_radius=12,
        padding=8,
        height=200,
    )

    page.add(
        top_bar,
        table_wrap,
        logs_wrap,
    )


if __name__ == "__main__":
    ft.app(target=main)
