from pyhwpx import Hwp
import win32com.client as win32
import pandas as pd
import re
import os


class EventTableProcessor:
    def __init__(self, hwp_path, excel_path):
        # HWP 문서를 열고 엑셀 파일을 읽음
        self.hwp = Hwp()
        self.hwp.Open(hwp_path)
        self.df = pd.read_excel(excel_path)
        self.date_counts = None  # 전처리된 날짜별 카운트를 저장

    # 처음  time column  필드값 전처리 함수
    def time_column_initial(self):
        # HWP 문서에서 시간 열 필드의 초기 상태 설정
        self.hwp.set_pos(15, 0, 0)
        self.hwp.TableCellBlock()  # 그냥 테이블 셀 선택
        self.hwp.TableCellBlockExtend()  # f5
        self.hwp.TableColPageDown()
        self.hwp.set_cur_field_name("")
        for i in range(7):
            self.hwp.set_pos(15 + i * 13, 0, 0)
            self.hwp.set_cur_field_name("time")
            # print(self.hwp.get_pos())

    # 근데 같은 월일안에서 TableLowerCell()을 실행하면 list값이 6씩증가하고
    # TableLowerCell()을 실행했을때 다른 월일 row로 넘어가면 +7이 된다

    # -------------------------#
    # 데이터 전처리

    def data_preprocessing(self):
        # 데이터를 전처리하여 삽입 준비
        self.df = self.df.replace(r'\n', '', regex=True)  # \n 제거
        self.df = self.df.fillna(" ")  # NaN 값은 공백으로 대체
        date_counts = self.df['date'].value_counts().reset_index()
        date_counts.columns = ['date', 'count']  # 열 이름 설정

        # 월과 일을 기준으로 날짜 정렬
        date_counts[['month', 'day']] = date_counts['date'].apply(
            lambda x: pd.Series(self.extract_month_day(x))
        )
        self.date_counts = date_counts.sort_values(by=['month', 'day']).drop(columns=['month', 'day']).reset_index(
            drop=True)
        # zprint(self.date_counts)

    def extract_month_day(self, date_str):
        # 날짜 문자열에서 월과 일을 추출
        match = re.search(r'(\d+)\.\s*(\d+)', date_str)
        if match:
            return int(match.group(1)), int(match.group(2))
        return None


    ##--------------표 생성 함수 -----------------------------##
    # 표 생성 함수 appendrow()

    def create_table(self):
        # 이벤트 수에 따라 테이블 행을 동적으로 생성
        for index, row in self.date_counts.iterrows():
            pset = self.hwp.HParameterSet.HTableDeleteLine
            self.hwp.move_to_field(f'time{{{{{index}}}}}')
            count = row['count']

            if count > 2:
                for i in range(count - 2):
                    self.hwp.TableAppendRow()
                    # print(f"{index}: append")
            elif count == 2:
                continue
            else:
                self.hwp.TableLowerCell()
                self.hwp.HAction.GetDefault("TableDeleteRow", pset.HSet)
                self.hwp.HAction.Execute("TableDeleteRow", pset.HSet)
                # print(f"{index}: delete")


#

# -----------------------------#-
# #데이터 넣는 함수

    def insert_data(self):
        # HWP 테이블에 날짜와 기타 데이터 삽입
        date_column_df = self.date_counts[['date']]

        # 날짜 값 삽입
        for index, row in date_column_df.iterrows():
            self.hwp.move_to_field(f'date{{{{{index}}}}}')
            self.hwp.insert_text(row['date'])
            # print(row['date'])

        # 날짜를 제외한 값 삽입
        df_no_date = self.df.drop(columns=['date'])
        self.hwp.set_pos(15, 0, 0)
        self.hwp.TableCellBlock()
        self.hwp.TableCellBlockExtend()
        self.hwp.TableColEnd()
        self.hwp.TableColPageDown()
        self.hwp.set_cur_field_name("imsi")
        self.hwp.put_field_text("imsi", df_no_date.values.flatten().tolist())

    import os

    def save_hwp_file(self, filename="output.hwp"):
        """
        HWP 파일을 지정된 위치에 저장하는 메서드.

        Parameters:
            filename (str): 저장할 파일 이름. 기본값은 "output.hwp".

        Returns:
            download_path (str): 파일이 저장된 경로.
        """
        # 사용자의 다운로드 폴더 경로 생성
        download_dir = os.path.join(os.path.expanduser("~"), 'downloads')
        base_name, ext = os.path.splitext(filename)
        download_path = os.path.join(download_dir, filename)

        # 동일한 파일명이 존재하는 경우 이름 변경 /  파일명(1) 이런식으로
        counter = 1
        while os.path.exists(download_path):
            new_filename = f"{base_name}({counter}){ext}"
            download_path = os.path.join(download_dir, new_filename)
            counter += 1

        # 파일 저장
        self.hwp.SaveAs(download_path)
        print(f"파일이 저장된 위치: {download_path}")
        return download_path

    def process(self):
        # 모든 작업을 순차적으로 실행
        self.time_column_initial()
        self.data_preprocessing()
        self.create_table()
        self.insert_data()
