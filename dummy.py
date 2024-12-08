import pandas as pd
import random
import os

# 폴더 생성
output_folder = "dummy_data"
os.makedirs(output_folder, exist_ok=True)

# 더미 데이터 생성
num_classes = 6
students_per_class = 23
gender_distribution = [10, 13]  # 남 10명, 여 13명
grades_distribution = {"A": 0.2, "B": 0.3, "C": 0.3, "D": 0.2}

# 각 반에 대한 데이터 생성
for class_num in range(1, num_classes + 1):
    data = []
    multicultural_added = False
    welfare_added = False

    for student_num in range(1, students_per_class + 1):
        student_name = f"예시{class_num}{student_num:02d}"

        # 성별 설정
        gender = "남" if student_num <= gender_distribution[0] else "여"

        # 성적 등급 설정 (무작위)
        grade = random.choices(list(grades_distribution.keys()), weights=grades_distribution.values())[0]

        # 생활지도 어려움 등급 설정 (무작위)
        difficulty_grade = random.choices(list(grades_distribution.keys()), weights=grades_distribution.values())[0]

        # 다문화 여부와 복지대상 여부 설정
        multicultural = "O" if not multicultural_added and student_num == random.randint(1, students_per_class) else "X"
        welfare = "O" if not welfare_added and student_num == random.randint(1, students_per_class) else "X"

        # 다문화 또는 복지대상 여부 추가 여부 설정
        if multicultural == "O":
            multicultural_added = True
        if welfare == "O":
            welfare_added = True

        data.append([student_name, gender, grade, difficulty_grade, multicultural, welfare])

    # 데이터프레임 생성
    df = pd.DataFrame(data, columns=["학생 이름", "성별", "성적 등급 (A/B/C/D)", "생활지도 어려움 등급 (A/B/C/D)", "다문화 여부 (O/X)", "복지대상 여부 (O/X)"])

    # 엑셀 파일로 저장
    output_path = os.path.join(output_folder, f"{class_num}.xlsx")
    df.to_excel(output_path, index=False)

output_folder
