from pythonHwp.hwpDeparmentCreate import EventTableProcessor

# processor = EventTableProcessor(hwp_path="../template2.hwp", excel_path="../eventdummy.xlsx")  # 인자값에 한글path랑 엑셀 패스 넣어줌
# processor.process() #한글 생성 함수
# processor.save_hwp_file("final_output.hwp")  # 파일을 "final_output.hwp" 이름으로 저장


# hwp_file = r"../template2.hwp"
# excel_file = r"../eventdummy.xlsx"
#
# event_processor = EventProcessor(hwp_file, excel_file)
# event_processor.process()



# 인스턴스 생성
processor = EventTableProcessor("../template2.hwp", "../eventdummy.xlsx")

# 모든 작업을 한 번에 처리
processor.process()

