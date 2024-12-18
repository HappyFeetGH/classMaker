import os, math
import pandas as pd
import random
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from collections import defaultdict


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

def assign_split_students(classes, split_data, data, num_classes):
    """
    분리 조건에 따라 학생을 각 학급에 무작위로 배치하는 함수.
    - split_data는 학번으로 구성됨.
    - 학생 데이터에는 "반"과 "번호"를 조합해 학번을 만들어 학생을 찾음.
    - 학급 순서를 무작위로 섞어서 특정 학급이 먼저 채워지는 것을 방지함.
    """
    # Step 1: 학급 순서를 무작위로 섞음
    class_names = list(classes.keys())
    random.shuffle(class_names)

    # Step 2: split_data의 학생 그룹도 무작위로 섞음
    split_groups = [group for group in split_data.values if not all(pd.isna(group))]
    random.shuffle(split_groups)

    # Step 3: 학생을 학번 기준으로 찾아 무작위 학급에 배치
    for group in split_groups:
        group = [int(student_id) for student_id in group if not pd.isna(student_id)]  # 유효한 학번만
        random.shuffle(group)  # 그룹 내부의 학생도 섞기

        for student_id in group:
            # 학번에서 반과 번호 추출
            target_class = student_id // 100  # 반 번호
            target_number = student_id % 100  # 번호

            # 학생 찾기
            student = next(
                (s for _, s in data.iterrows() if s["반"] == target_class and s["번호"] == target_number),
                None
            )

            if student is not None:
                # 학급을 무작위 순서대로 돌아가며 배치
                for current_class in class_names:
                    if len(classes[current_class]) < sum(len(classes[c]) for c in class_names) // num_classes + 1:
                        classes[current_class].append(student.to_dict())
                        data = data[(data["반"] != target_class) | (data["번호"] != target_number)]
                        break
    return data


def balance_by_gender_and_count(data, classes, num_classes, weights, max_per_class_male=2, max_per_class_female=3):
    """
    남녀 비율, 특정 반 출신 학생 수, 추가 점수 조건을 고려하여 학생을 배치.
    배치할 수 없는 학생은 별도 Class_X로 배치.
    """
    # Step 1: 학생 데이터를 섞기
    students = data.to_dict("records")
    random.shuffle(students)

    # Step 2: 학급별 배치 상태 초기화
    class_names = list(classes.keys())
    class_prev_counts = {class_name: defaultdict(int) for class_name in class_names}
    class_gender_counts = {class_name: {"남": 0, "여": 0} for class_name in class_names}

    # Step 3: 배치 불가능한 학생 저장용
    unassigned_students = []

    def calculate_class_score(current_class, student):
        """
        학급의 점수를 계산:
        1. 특정 반 출신 학생 수
        2. 남녀 불균형 점수
        3. 추가 조건 점수 (성적 등급, 생활지도 어려움 등급 등)
        """
        prev_class = student["반"]
        gender = student["성별"]
        student_score = calculate_score(student, weights)

        # 특정 반 출신 점수 (출신 반 학생 수가 많을수록 패널티)
        prev_class_penalty = class_prev_counts[current_class][prev_class]

        # 남녀 불균형 점수
        gender_balance_penalty = abs(
            (class_gender_counts[current_class]["남"] + (1 if gender == "남" else 0)) -
            (class_gender_counts[current_class]["여"] + (1 if gender == "여" else 0))
        )

        # 전체 점수 = 반 출신 패널티 + 남녀 불균형 + 학생의 개별 점수
        total_score = prev_class_penalty + gender_balance_penalty + student_score
        return total_score

    # Step 4: 학생 배치 (그리디 + 제약 조건)
    for student in students:
        prev_class = student["반"]
        gender = student["성별"]

        # 각 학급의 점수 계산
        candidate_classes = []
        for current_class in class_names:
            score = calculate_class_score(current_class, student)
            total_students = len(classes[current_class])
            candidate_classes.append((current_class, score, total_students))

        # 학급을 점수(출신 학생 수 + 남녀 균형 + 추가 조건) 및 총원 기준으로 정렬
        candidate_classes.sort(key=lambda x: (x[1], x[2]))

        # 가장 균형 잡힌 학급에 배치
        assigned = False
        for current_class, _, _ in candidate_classes:
            if gender == "남" and class_prev_counts[current_class][prev_class] < max_per_class_male:
                classes[current_class].append(student)
                class_prev_counts[current_class][prev_class] += 1
                class_gender_counts[current_class][gender] += 1
                assigned = True
                break
            elif gender == "여" and class_prev_counts[current_class][prev_class] < max_per_class_female:
                classes[current_class].append(student)
                class_prev_counts[current_class][prev_class] += 1
                class_gender_counts[current_class][gender] += 1
                assigned = True
                break

        if not assigned:
            unassigned_students.append(student)  # 배치할 수 없는 학생 저장

    # Step 5: 배치 불가능한 학생을 Class_X에 추가
    classes["Class_X"] = unassigned_students

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
    global weights
    # 학급 초기화
    classes = {f"Class_{i+1}": [] for i in range(num_classes)}

    # 데이터 로드
    data = load_data_from_folder(input_folder)
    split_data = pd.read_excel(split_file_path, header=None)

    # Step 1: 분리 조건 학생 배치
    remaining_data = assign_split_students(classes, split_data, data, num_classes)

    # Step 2: 남/여 및 학급별 학생 수 맞춰 배치
    #balance_by_gender_and_count(remaining_data, classes, num_classes)
    classes = balance_by_gender_and_count(remaining_data, classes, num_classes, weights, 3, 4)

    # Step 3: 조건 기반 밸런싱
    #weights = {
    #    "성적 등급 (A/B/C/D)": 1,
    #    "생활지도 어려움 등급 (A/B/C/D)": 2
    #}
    #swap_students_between_classes(classes, weights, split_data)
    
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
        "성적 등급 (A/B/C/D)": 5,
        "생활지도 어려움 등급 (A/B/C/D)": 3,
        "체력 (A/B/C/D)": 1
}

# 실행
num_classes = 5  # 원하는 반의 수
classes = assign_classes(num_classes)
save_results(classes)

