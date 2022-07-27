import re
from argparse import ArgumentParser
from pathlib import PurePath

import os
import csv
import xlsxwriter

from typing import List


def get_args():
    current_workdir = os.getcwd()

    arg_parser = ArgumentParser()
    arg_parser.add_argument("-r", "--report", default=f"{str(PurePath(current_workdir, 'report.csv'))}")
    arg_parser.add_argument("-t", "--theater", default=f"{str(PurePath(current_workdir, 'theater.csv'))}")

    return arg_parser.parse_args()


def read_csv(path: PurePath) -> List[List[str]]:
    csv_data: List[List[str]] = []
    with open(path, newline='', encoding='utf-8') as csv_file:
        csv_reader = csv.reader(csv_file)
        for row in csv_reader:
            csv_data.append(row)

    return csv_data


def prep_report(data: List[List[str]]) -> List[List[str]]:
    in_data_area = False
    start_idx_data = -1
    end_idx_data = -1

    for idx, row in enumerate(data):
        if len(row) == 4 and row[0] == "Rec" and row[1] == "Bestellnummer" and row[2] == "Artikeldetails":
            start_idx_data = idx + 1
            in_data_area = True
        if in_data_area and len(row) == 1 and row[0] == "":
            end_idx_data = idx
            break

    sanitized_data: List[List[str]] = []
    for row in data[start_idx_data:end_idx_data]:
        section_list = row[2].split("-", maxsplit=1)
        type_place_list: List[str] = section_list[1].rsplit("-", maxsplit=1)

        ticket_type = type_place_list[0].strip()

        pattern = re.compile(r"Row:\s*(?P<row>\d+)\s*Seat:\s*(?P<seat>\d+)", re.VERBOSE)
        match = pattern.match(type_place_list[1].strip())
        row_place = int(match.group("row"))
        seat = int(match.group("seat"))

        sanitized_data.append([row[0], row[1], ticket_type, f"{row_place}-{seat:02d}"])

    return sanitized_data


def prep_theater(data: List[List[str]]) -> List[List[str]]:
    _, *tail = data

    sanitized_data: List[List[str]] = []
    for row in tail:
        row_data = [row[3], row[28], row[37]]
        sanitized_data.append(row_data)

    return sanitized_data


def find_matching_dataset(t: List[List[str]], orderId: str, abo_type: str) -> int:
    for idx, data in enumerate(t):
        if data[0] == orderId and abo_type in data[1]:
            return idx
    exit(999)


def merge_data(rep: List[List[str]], the: List[List[str]]) -> List[List[str]]:
    theater_list = the.copy()

    data = []

    for report_row in rep:
        theater_idx = find_matching_dataset(theater_list, report_row[1], report_row[2])
        sub_code = theater_list[theater_idx][2]
        del theater_list[theater_idx]

        data.append([*report_row, sub_code])

    return data


if __name__ == "__main__":
    # args.report
    # args.theater
    args = get_args()

    report = read_csv(args.report)
    report_sanitized = prep_report(report)

    theater = read_csv(args.theater)
    theater_sanitized = prep_theater(theater)

    merged = merge_data(report_sanitized, theater_sanitized)

    data_dict = {}

    for wh_row in merged:
        if wh_row[2] not in data_dict.keys():
            data_dict[wh_row[2]] = []
        data_dict[wh_row[2]].append(wh_row)


    for key in data_dict.keys():
        data = data_dict[key]
        with_header = [["Lfd.-Nr.", "Buchungsnummer", "Ticket-Typ", "Platz", "Abocode"], *data]

        workbook = xlsxwriter.Workbook(f"{key}.xlsx")
        worksheet = workbook.add_worksheet()

        for idx, wh_row in enumerate(with_header):
            for idy, wh_cell in enumerate(wh_row):
                worksheet.write(idx, idy, wh_cell)

        workbook.close()






