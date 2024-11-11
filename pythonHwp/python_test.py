from pythonHwp.hwpDeparmentCreate import EventTableProcessor

#파일 save할떄 만약 이미 같은 경로에 파일이 있다면 warning이라든지 클라에게 표시를 해줘야함

processor = EventTableProcessor(hwp_path="../template2.hwp", excel_path="../eventdummy.xlsx")  # 인자값에 한글path랑 엑셀 패스 넣어줌
processor.process() #한글 생성 함수
processor.save_hwp_file("final_output.hwp")  # 파일을 "final_output.hwp" 이름으로 저장
