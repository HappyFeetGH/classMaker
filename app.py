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
    output_file_path = "./xlsx_storage/Class_Assignment_Results.xlsx"
    with pd.ExcelWriter(output_file_path) as writer:
        for class_name, df in final_classes.items():
            df.to_excel(writer, sheet_name=class_name, index=False)
    return jsonify({"message": "Excel file saved successfully"})

if __name__ == '__main__':
    app.run(debug=True)
