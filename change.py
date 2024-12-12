import pandas as pd
import os

folder_path = "./xlsx_storage"  # 원본 파일 경로
for file in os.listdir(folder_path):
    # 기존 데이터 로드
    file_path = os.path.join(folder_path, file)
    data = pd.read_excel(file_path)

    # 반과 번호 컬럼 추가
    data["반"] = data["학생 이름"].str.extract(r"(\d)").astype(int)  # '학생 이름'에서 반 추출
    data["번호"] = [x+1 for x in range(len(data["반"]))]

    # 컬럼 순서 재정렬
    columns = ["반", "번호"] + [col for col in data.columns if col not in ["반", "번호"]]
    data = data[columns]

    # 수정된 데이터 저장
    output_file_path = file_path
    data.to_excel(output_file_path, index=False)

    
