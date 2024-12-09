import os, math
import pandas as pd
import random
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# 경로 설정
input_folder = "./xlsx_storage"
output_file_path = "Class_Assignment_Results.xlsx"
split_file_path = "split_them.xlsx"

# 파일 읽기 함수
def load_data_from_folder(folder_path):
    dataframes = []
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx"):
            file_path = os.path.join(folder_path, file_name)
            df = pd.read_excel(file_path)
            # 전 학년 반 정보 추가
            df["전 학년 반"] = file_name.split(".")[0]  # 파일 이름에서 반 번호 추출
            dataframes.append(df)
    return pd.concat(dataframes, ignore_index=True)
# 반별 MAX/MIN 계산 함수
def calculate_limits(total, num_classes):
    return math.ceil(total / num_classes), math.floor(total / num_classes)

# 분리 조건 우선 배치 함수
def assign_split_students(classes, split_data, data, num_classes):
    for group in split_data.values:
        group = [name for name in group if not pd.isna(name)]  # 유효한 이름만
        for i, student_name in enumerate(group):
            student = next((s for _, s in data.iterrows() if s["학생 이름"] == student_name), None)
            if student is not None:
                class_index = i % num_classes
                classes[f"Class_{class_index + 1}"].append(student.to_dict())
                data = data[data["학생 이름"] != student_name]  # 배치된 학생 제거
    return data

# 남/여 및 학급별 학생 수 균형 있게 배치 (학생 섞기 + 학급 랜덤 배치)
def balance_by_gender_and_count(data, classes, num_classes):
    # 남학생과 여학생 데이터 분리
    male_students = data[data["성별"] == "남"].to_dict("records")
    female_students = data[data["성별"] == "여"].to_dict("records")

    # 학생 데이터 섞기
    random.shuffle(male_students)
    random.shuffle(female_students)

    # Step 1: 전체 남학생/여학생을 학급당 나눠야 할 수 계산
    total_male_students = len(male_students) + sum(
        sum(1 for s in class_students if s["성별"] == "남") for class_students in classes.values()
    )
    total_female_students = len(female_students) + sum(
        sum(1 for s in class_students if s["성별"] == "여") for class_students in classes.values()
    )

    males_per_class = [total_male_students // num_classes] * num_classes
    females_per_class = [total_female_students // num_classes] * num_classes

    # 나머지 남학생/여학생을 앞 학급부터 채우기
    for i in range(total_male_students % num_classes):
        males_per_class[i] += 1
    for i in range(total_female_students % num_classes):
        females_per_class[i] += 1

    # Step 2: 이미 배치된 남학생/여학생 수 계산
    current_male_counts = {
        class_name: sum(1 for s in class_students if s["성별"] == "남") for class_name, class_students in classes.items()
    }
    current_female_counts = {
        class_name: sum(1 for s in class_students if s["성별"] == "여") for class_name, class_students in classes.items()
    }

    # 각 학급에 추가로 배치해야 할 남/여 학생 수 계산
    remaining_males_needed = {
        class_name: max(0, males_per_class[i] - current_male_counts[class_name])
        for i, class_name in enumerate(classes.keys())
    }
    remaining_females_needed = {
        class_name: max(0, females_per_class[i] - current_female_counts[class_name])
        for i, class_name in enumerate(classes.keys())
    }

    # Step 3: 학급에 남학생 배치
    male_index = 0
    for class_name in random.sample(list(classes.keys()), len(classes)):  # 학급 순서 무작위
        needed_males = remaining_males_needed[class_name]
        for _ in range(needed_males):
            if male_index < len(male_students):
                classes[class_name].append(male_students[male_index])
                male_index += 1

    # Step 4: 학급에 여학생 배치
    female_index = 0
    for class_name in random.sample(list(classes.keys()), len(classes)):  # 학급 순서 무작위
        needed_females = remaining_females_needed[class_name]
        for _ in range(needed_females):
            if female_index < len(female_students):
                classes[class_name].append(female_students[female_index])
                female_index += 1

    return classes




# 가중치 계산 함수
def calculate_score(student, weights):
    score = 0
    for key, weight in weights.items():
        score += weight * (ord(student[key]) - ord('A') + 1)  # A=1, B=2, ...
    return score


def violates_split_condition(student_name, recipient_class, split_data):
    for group in split_data.values:
        if student_name in group:
            # 같은 그룹에 속한 학생이 recipient_class에 있는지 확인
            for existing_student in recipient_class:
                if existing_student["학생 이름"] in group:
                    return True
    return False

# 학급 조건 확인 함수
def is_class_balanced(classes, male_max, male_min, female_max, female_min):
    for class_key, class_students in classes.items():
        df = pd.DataFrame(class_students)
        male_count = df["성별"].value_counts().get("남", 0)
        female_count = df["성별"].value_counts().get("여", 0)
        if not (male_min <= male_count <= male_max and female_min <= female_count <= female_max):
            return False
    return True

# 학생 교환 함수
def swap_students_between_classes(classes, weights, split_data, max_iterations=100):
    iteration_count = 0

    while iteration_count < max_iterations:
        swapped = False
        for class_a in classes:
            for class_b in classes:
                if class_a == class_b:
                    continue

                # 각 학급의 학생 데이터
                students_a = classes[class_a]
                students_b = classes[class_b]

                # 교환 가능한 학생 쌍 찾기
                for student_a in students_a:
                    for student_b in students_b:
                        # 성별이 동일하고 분리 조건 위반하지 않는 경우
                        if student_a["성별"] == student_b["성별"] and not violates_split_condition(student_a["학생 이름"], students_b, split_data) and not violates_split_condition(student_b["학생 이름"], students_a, split_data):
                            # 교환 후 점수 계산
                            original_score_a = calculate_class_score(students_a, weights)
                            original_score_b = calculate_class_score(students_b, weights)

                            # 교환
                            temp_a = students_a.copy()
                            temp_b = students_b.copy()

                            temp_a.remove(student_a)
                            temp_a.append(student_b)
                            temp_b.remove(student_b)
                            temp_b.append(student_a)

                            new_score_a = calculate_class_score(temp_a, weights)
                            new_score_b = calculate_class_score(temp_b, weights)

                            # 점수 균형 개선 시 교환 실행
                            if abs(new_score_a - new_score_b) < abs(original_score_a - original_score_b):
                                students_a.remove(student_a)
                                students_a.append(student_b)
                                students_b.remove(student_b)
                                students_b.append(student_a)
                                swapped = True
                                break
                    if swapped:
                        break
            if swapped:
                break

        if not swapped:
            break  # 더 이상 교환이 불가능하면 종료
        iteration_count += 1

    if iteration_count >= max_iterations:
        print("Warning: Swap iterations exceeded limit.")
    else:
        print(f"Student swaps completed in {iteration_count} iterations.")

# 학급 점수 계산 함수
def calculate_class_score(class_students, weights):
    return sum(calculate_score(student, weights) for student in class_students) / len(class_students)


# 요약 정보 생성
def generate_summary(classes, weights):
    summary_data = []
    for class_name, class_students in classes.items():
        df = pd.DataFrame(class_students)

        # 학급 점수 계산
        total_score = sum(calculate_score(student, weights) for student in class_students)

        summary = {
            "반": class_name,
            "총 학생 수": len(class_students),
            "남학생 수": df["성별"].value_counts().get("남", 0),
            "여학생 수": df["성별"].value_counts().get("여", 0),
            "학급 총점": total_score,
        }
        for grade in ["A", "B", "C", "D"]:
            summary[f"성적 {grade}"] = df["성적 등급 (A/B/C/D)"].value_counts().get(grade, 0)
            summary[f"생활지도 어려움 {grade}"] = df["생활지도 어려움 등급 (A/B/C/D)"].value_counts().get(grade, 0)
        summary_data.append(summary)
    return pd.DataFrame(summary_data)

# 학급 점수 계산 함수
def calculate_score(student, weights):
    score = 0
    for key, weight in weights.items():
        if key in student and student[key] in "ABCD":
            score += weight * (ord(student[key]) - ord('A') + 1)  # A=1, B=2, C=3, D=4
    return score

# 동적 반 편성 함수
def assign_classes(num_classes):
    # 학급 초기화
    classes = {f"Class_{i+1}": [] for i in range(num_classes)}

    # 데이터 로드
    data = load_data_from_folder(input_folder)
    split_data = pd.read_excel(split_file_path, header=None)

    # Step 1: 분리 조건 학생 배치
    remaining_data = assign_split_students(classes, split_data, data, num_classes)

    # Step 2: 남/여 및 학급별 학생 수 맞춰 배치
    balance_by_gender_and_count(remaining_data, classes, num_classes)

    # Step 3: 조건 기반 밸런싱
    weights = {
        "성적 등급 (A/B/C/D)": 1,
        "생활지도 어려움 등급 (A/B/C/D)": 2
    }
    swap_students_between_classes(classes, weights, split_data)

    return classes

final_classes = {}  # 전역 변수로 선언


# 결과 저장 및 시각화
def save_results(classes):
    global final_classes
    # 결과를 학급별로 DataFrame으로 변환
    final_classes = {
        class_name: pd.DataFrame(class_students).fillna("")
        for class_name, class_students in classes.items()
    }

    # Excel 파일 저장
    with pd.ExcelWriter(output_file_path) as writer:
        for class_name, df in final_classes.items():
            df.to_excel(writer, sheet_name=class_name, index=False)

        summary_df = generate_summary(classes, weights)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

    print(f"Class assignment with summaries saved to {output_file_path}")

weights = {
        "성적 등급 (A/B/C/D)": 1,
        "생활지도 어려움 등급 (A/B/C/D)": 2
}

# 실행
num_classes = 6  # 원하는 반의 수
classes = assign_classes(num_classes)
save_results(classes)

