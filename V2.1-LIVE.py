
import os
from pathlib import Path
import re
import warnings
import json
import pandas as pd
import openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

username = os.getlogin()
username_file_path = os.path.join(os.path.expanduser('~'))
download_path = f"{username_file_path}\\Downloads\\"

file = {
    "sh_file_path": f"{download_path}Smart Hands Details.xlsx",
    "tt_file_path": f"{download_path}Trouble Ticket Charges Report.xlsx",  # legacy fallback
    "labour_file_path": f"{download_path}Labour.xlsx",
    "commit_file_path": f"{download_path}Commit.xlsx",
    "activity_file_path": f"{download_path}Order Activities.xlsx",
    "ssc_report_file_path": f"{download_path}SSC - Management Reporting Data 2024.xlsx",
}

class Report:
    def __init__(self):
        self.sh_file_path = file['sh_file_path']
        self.tt_file_path = file['tt_file_path']
        self.labour_file_path = file['labour_file_path']
        self.commit_file_path = file['commit_file_path']
        self.activity_file_path = file['activity_file_path']
        self.ssc_report_file_path = file['ssc_report_file_path']
        self.download_path = download_path

    def labour(self) -> list:
        labour = pd.read_excel(self.labour_file_path)
        labour_filtered = labour[(labour['User Total Hours'] > 0) | (labour['User Total Minutes'] > 0)]
        miss_count = len(labour[labour['User Total Hours'] == 0])
        ncc_count = len(labour)
        labour_cap_percent = round(((ncc_count - miss_count) / ncc_count) * 100, 2) if ncc_count else 0
        labour_cap_hours = round(labour_filtered['User Total Hours'].sum(), 2)
        return labour_cap_percent, labour_cap_hours

    def smart_hand(self) -> list:
        sh = pd.read_excel(self.sh_file_path)
        sh_size = sh.shape
        billable_flag = [sh.iloc[k, 11] for k in range(0, sh_size[0] - 2)]
        yescount = len([flag for flag in billable_flag if flag == 'Y'])
        nocount = len([flag for flag in billable_flag if flag != 'Y'])
        sh_hour = sh.iloc[:, 6].sum()
        sh_minute = sh.iloc[:, 7].sum()
        return yescount, nocount, sh_hour, sh_minute

    def trouble_ticket(self) -> list:
        tt1 = pd.read_excel(os.path.join(self.download_path, "trouble1.xlsx"))
        tt2 = pd.read_excel(os.path.join(self.download_path, "trouble2.xlsx"))

        # Drop empty columns to prevent concat warning
        tt1 = tt1.dropna(axis=1, how='all')
        tt2 = tt2.dropna(axis=1, how='all')
        tt = pd.concat([tt1, tt2], ignore_index=True)

        # Fill NaNs only on non-datetime columns
        for col in tt.select_dtypes(exclude=['datetime']).columns:
            tt[col] = tt[col].fillna(0)

        severity_counts = tt['Ticket Severity'].value_counts()
        lowcount = severity_counts.get('Low', 0)
        mediumcount = severity_counts.get('Medium', 0)
        highcount = severity_counts.get('High', 0)
        tt_hour = pd.to_numeric(tt['Labor Hours Billed'], errors='coerce').fillna(0).sum()
        tt_minute = pd.to_numeric(tt['Labor Mins Billed'], errors='coerce').fillna(0).sum()
        return lowcount, mediumcount, highcount, tt_hour, tt_minute

    def commit(self):
        commit = pd.read_excel(self.commit_file_path)
        ssc_report_stat = pd.read_excel(self.ssc_report_file_path)
        commit_size = commit.shape
        total_activities = commit_size[0] - 2
        total_missed_list = [flag for flag in commit.iloc[:, 22] if flag == 'Missed']
        commit_miss_count = len(total_missed_list)
        filtered_frame = commit[commit['Compliance Status'] == "Missed"]
        filtered_frame = filtered_frame[['Compliance Status', 'Commit Date Missed Category']]
        categories = filtered_frame['Commit Date Missed Category']
        non_ops_caused_list = categories[categories == 'Equinix Non-Ops Caused'].tolist()
        customer_caused_list = categories[categories == 'Customer Caused'].tolist()
        system_caused_list = categories[categories == 'System Caused'].tolist()
        ops_caused_list = categories[categories == 'Ops Caused'].tolist()
        blank_list = categories[categories == 'Blank'].tolist()
        on_time_activities = total_activities - commit_miss_count
        summary = pd.read_excel(file['ssc_report_file_path'], 'Commit Compliance')
        ytd_commit_rate = round(float(summary.iloc[53, 6]) * 100, 2)
        commit_rate = round((on_time_activities / total_activities) * 100, 2) if total_activities > 0 else 0
        return total_activities, on_time_activities, commit_miss_count, commit_rate, ytd_commit_rate, customer_caused_list, non_ops_caused_list, system_caused_list, ops_caused_list, blank_list

    def activity(self):
        df = pd.read_excel(self.activity_file_path).iloc[1:]
        install = df[df['Unnamed: 35'] == "Install"].shape[0]
        de_install = df[df['Unnamed: 35'] == "De-Install"].shape[0]
        return install, de_install
