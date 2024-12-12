from flask import Flask, jsonify, request, render_template
from main import final_classes, calculate_score, generate_summary, weights  # main.py에서 함수 임포트
import pandas as pd

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/get_classes', methods=['GET'])
def get_classes():
    # 학급별 요약 정보 계산
    json_classes = {}
    #weights = {"성적 등급 (A/B/C/D)": 1, "생활지도 어려움 등급 (A/B/C/D)": 2}  # 가중치 설정

    for class_name, df in final_classes.items():
        df = df.fillna("")  # NaN 값 처리

        # 학급 요약 정보
        total_students = int(len(df))
        male_students = int(len(df[df["성별"] == "남"]))
        female_students = int(len(df[df["성별"] == "여"]))
        total_score = int(df.apply(lambda row: calculate_score(row, weights), axis=1).sum())

        # DataFrame을 dict로 변환 (int64 -> int 변환)
        json_classes[class_name] = {
            "students": df.astype({"성별": "str"}).to_dict(orient="records"),  # 모든 데이터를 JSON 호환 타입으로 변환
            "summary": {
                "총 학생 수": total_students,
                "남학생 수": male_students,
                "여학생 수": female_students,
                "학급 전체 점수": total_score,
            }
        }

    return jsonify(json_classes)

@app.route('/update_classes', methods=['POST'])
def update_classes():
    global final_classes
    updated_data = request.json  # 프론트엔드에서 보낸 수정 데이터

    try:
        # 데이터 구조에 맞게 "students" 키를 접근
        final_classes = {
            class_name: pd.DataFrame(class_data["students"])
            for class_name, class_data in updated_data.items()
        }
    except Exception as e:
        print("Error in data structure:", updated_data)  # 디버깅을 위해 데이터 출력
        return jsonify({"error": f"Invalid data structure: {str(e)}"}), 400

    return jsonify({"message": "Classes updated successfully"})


@app.route('/save_to_excel', methods=['POST'])
def save_to_excel():
    global final_classes

    # 데이터 검증
    if not final_classes:
        return jsonify({"error": "No class data to save"}), 400

    #try:
    # **원본 반별 데이터 저장**
    # **새로운 학급 이름 계산**
    for i, class_name in enumerate(final_classes.keys(), start=1):
        class_letter = chr(64 + i)  # 'class_1' -> A, 'class_2' -> B ...
        final_classes[class_name]["새로운 학급"] = class_letter

    # **원본 반별 데이터 저장**
    original_output_path = "./Original_Class_Data.xlsx"
    with pd.ExcelWriter(original_output_path) as writer:
        # 모든 학생 데이터를 통합
        all_students = pd.concat(final_classes.values(), ignore_index=True)

        # 원본 반별로 순회
        for original_class in sorted(all_students["반"].unique()):
            # 해당 반에 속한 학생 필터링
            class_data = all_students[all_students["반"] == original_class]

            # **번호 순으로 정렬**
            class_data = class_data.sort_values(by=["번호"])

            # 남학생/여학생 나누기
            males = class_data[class_data["성별"] == "남"]
            females = class_data[class_data["성별"] == "여"]

            # 시트 이름: 원본 반 번호
            sheet_name = f"{original_class}반"

            # 남학생 데이터 저장
            males_summary = males[["번호", "학생 이름", "새로운 학급", "비고"]]
            males_summary.rename(columns={"학생 이름": "남학생 이름"}, inplace=True)
            males_summary.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)

            # 여학생 데이터 저장
            females_summary = females[["번호", "학생 이름", "새로운 학급", "비고"]]
            females_summary.rename(columns={"학생 이름": "여학생 이름"}, inplace=True)
            females_summary.to_excel(writer, sheet_name=sheet_name, index=False, startrow=len(males_summary) + 3)

    # **새로운 반별 데이터 저장**
    new_output_path = "./New_Class_Data.xlsx"
    with pd.ExcelWriter(new_output_path) as writer:
        for class_number, class_data in final_classes.items():
            if not isinstance(class_data, pd.DataFrame):
                continue

            # **새로운 학급 이름 추가 (A, B, C, ...)**
            class_letter = chr(64 + int(class_number.split('_')[1]))  # 'class_1' -> A, 'class_2' -> B ...
            class_data["새로운 학급"] = class_letter

            # **이름 순으로 정렬**
            class_data = class_data.sort_values(by=["학생 이름"])

            # 남학생/여학생 나누기
            males = class_data[class_data["성별"] == "남"]
            females = class_data[class_data["성별"] == "여"]

            # 시트 이름: 새로운 반 이름 (A반, B반, ...)
            sheet_name = f"{class_letter}반"

            # 남학생 데이터 저장
            males_summary = males[["학생 이름", "반", "비고"]]
            males_summary.rename(columns={"학생 이름": "남학생 이름", "반": "이전 학급"}, inplace=True)
            males_summary.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)

            # 여학생 데이터 저장
            females_summary = females[["학생 이름", "반", "비고"]]
            females_summary.rename(columns={"학생 이름": "여학생 이름", "반": "이전 학급"}, inplace=True)
            females_summary.to_excel(writer, sheet_name=sheet_name, index=False, startrow=len(males_summary) + 3)

        # 성공 메시지 반환
        return jsonify({
            "message": "Excel files saved successfully",
            "original_file_path": original_output_path,
            "new_file_path": new_output_path
        })

    #except Exception as e:
    #    print(f"Error saving Excel file: {e}")  # 디버깅 로그
    #    return jsonify({"error": "Failed to save Excel files"}), 500







if __name__ == '__main__':
    #app.run(debug=True)
    app.run(host='0.0.0.0', port=5000, debug=True)
