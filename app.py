from flask import Flask, jsonify, request, render_template, send_file
from main import final_classes, calculate_score, generate_summary, weights, split_data  # main.py에서 함수 임포트
import pandas as pd
import io
import zipfile

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

@app.route('/get_split_data', methods=['GET'])
def get_split_data():
    try:
        split_data_list = []
        for _, row in split_data.iterrows():
            group = [int(num) for num in row if not pd.isna(num)]  # NaN 제거
            split_data_list.append(group)
        return jsonify(split_data_list)
    except Exception as e:
        print(f"Error processing split data: {e}")
        return jsonify({"error": "Failed to process split data"}), 500


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


@app.route('/download_excel', methods=['GET'])
def download_excel():
    global final_classes

    # 데이터 검증
    if not final_classes:
        return jsonify({"error": "No class data to download"}), 400

    # 새로운 학급 이름 추가
    for i, class_name in enumerate(final_classes.keys(), start=1):
        class_letter = chr(64 + i)  # 'class_1' -> A, 'class_2' -> B ...
        if not class_name.startswith("Class_X"):  # Class_X는 배제
            final_classes[class_name]["새로운 학급"] = class_letter

    # 미배정 학생 처리: Class_X가 없으면 빈 데이터프레임 추가
    if "Class_X" not in final_classes:
        final_classes["Class_X"] = pd.DataFrame(columns=["학생 이름", "반", "번호", "성별", "비고"])

    try:
        # 메모리에 ZIP 파일 생성
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            # **1. Original_Class_Data.xlsx 생성**
            original_output = io.BytesIO()
            with pd.ExcelWriter(original_output, engine='openpyxl') as writer:
                # 모든 학생 데이터를 통합
                all_students = pd.concat(final_classes.values(), ignore_index=True)

                # 원본 반별 데이터 저장
                for original_class in sorted(all_students["반"].unique()):
                    # 해당 반에 속한 학생 필터링
                    class_data = all_students[all_students["반"] == original_class]

                    # 번호 순으로 정렬
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

                # 미배정 학생(Class_X) 저장
                unassigned = final_classes["Class_X"]  # 빈 데이터프레임이어도 처리 가능
                unassigned.to_excel(writer, sheet_name="미배정 학생", index=False)

            # ZIP 파일에 추가
            original_output.seek(0)
            zip_file.writestr("Original_Class_Data.xlsx", original_output.read())

            # **2. New_Class_Data.xlsx 생성**
            new_output = io.BytesIO()
            with pd.ExcelWriter(new_output, engine='openpyxl') as writer:
                for class_number, class_data in final_classes.items():
                    if not isinstance(class_data, pd.DataFrame):
                        continue

                    # **빈 학급 데이터 확인**
                    if class_data.empty:
                        print(f"Skipping empty class: {class_number}")
                        continue

                    # **필수 열 확인 및 기본값 추가**
                    required_columns = ["학생 이름", "반", "비고"]
                    for col in required_columns:
                        if col not in class_data.columns:
                            print(f"Missing column '{col}' in class {class_number}. Adding default.")
                            class_data[col] = ""  # 기본값 추가

                    # **새로운 학급 이름 추가 (A, B, C, ...)**
                    if not class_number.startswith("Class_X"):
                        class_letter = chr(64 + int(class_number.split('_')[1]))
                        class_data["새로운 학급"] = class_letter

                    # **이름 순으로 정렬**
                    try:
                        class_data = class_data.sort_values(by=["학생 이름"])
                    except KeyError as e:
                        print(f"Error sorting class {class_number}: {e}")
                        continue

                    # 남학생/여학생 나누기
                    males = class_data[class_data["성별"] == "남"]
                    females = class_data[class_data["성별"] == "여"]

                    # 시트 이름: 새로운 반 이름 (A반, B반, ...)
                    sheet_name = f"{class_letter}반" if not class_number.startswith("Class_X") else "미배정 학생"

                    # 남학생 데이터 저장
                    try:
                        males_summary = males[["학생 이름", "반", "비고"]]
                        males_summary.rename(columns={"학생 이름": "남학생 이름", "반": "이전 학급"}, inplace=True)
                        males_summary.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
                    except KeyError as e:
                        print(f"Error processing male data for class {class_number}: {e}")
                        continue

                    # 여학생 데이터 저장
                    try:
                        females_summary = females[["학생 이름", "반", "비고"]]
                        females_summary.rename(columns={"학생 이름": "여학생 이름", "반": "이전 학급"}, inplace=True)
                        females_summary.to_excel(writer, sheet_name=sheet_name, index=False, startrow=len(males_summary) + 3)
                    except KeyError as e:
                        print(f"Error processing female data for class {class_number}: {e}")
                        continue

            # ZIP 파일에 추가
            new_output.seek(0)
            zip_file.writestr("New_Class_Data.xlsx", new_output.read())


        # ZIP 파일 반환
        zip_buffer.seek(0)
        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name="Class_Data.zip",
            mimetype="application/zip"
        )

    except Exception as e:
        print(f"Error creating ZIP file: {e}")  # 디버깅 로그
        return jsonify({"error": "Failed to create ZIP file"}), 500




if __name__ == '__main__':
    #app.run(debug=True)
    app.run(host='0.0.0.0', port=5000, debug=True)
