let classesData = {};
let numClasses = 0;
let currentSplitData = [];

// Fetch initial data
async function fetchClasses() {
    const response = await fetch('/get_classes');
    classesData = await response.json();
    renderClasses();
}

// Render classes
function renderClasses() {
    const { totalMaleStudents, totalFemaleStudents } = calculateTotalStudents(classesData);
    const container = document.getElementById('classes-container');
    container.innerHTML = '';

    for (const [className, classData] of Object.entries(classesData)) {
        const classBox = document.createElement('div');
        classBox.classList.add('class-box');
        classBox.setAttribute('data-class-name', className); // 학급 이름 저장
        classBox.addEventListener('dragover', allowDrop);
        classBox.addEventListener('drop', handleDrop);

        // 학급 이름 추가
        const classHeader = document.createElement('h3');
        classHeader.textContent = className === "Class_X" ? "미배정 학생" : className;
        classBox.appendChild(classHeader);

        // 학급 요약 정보 추가 (Class_X는 요약 제외)
        if (className !== "Class_X") {
            const summary = document.createElement('div');
            summary.classList.add('class-summary');
            summary.innerHTML = `
                ${renderSummaryLine("남학생 수", classData.summary["남학생 수"], totalMaleStudents, numClasses)}
                ${renderSummaryLine("여학생 수", classData.summary["여학생 수"], totalFemaleStudents, numClasses)}
                ${renderSummaryLine("총 학생 수", classData.summary["총 학생 수"], totalMaleStudents + totalFemaleStudents, numClasses)}
            `;
            classBox.appendChild(summary);
        }

        
        // 학생 목록 추가
        classData.students.forEach((student, index) => {
            const studentDiv = document.createElement('div');
            studentDiv.classList.add('student');
            studentDiv.setAttribute('data-student-id', index); // 학생 인덱스 저장
            studentDiv.setAttribute('draggable', true);
            //studentDiv.textContent = `${student["학생 이름"]} (${student["성별"]}) - ${student["비고"]}`;
            studentDiv.setAttribute("data-origin-class", student["전 학년 반"]); // 출신 반 저장
            studentDiv.innerHTML = `
                <div>${student["학생 이름"]} (${student["성별"]}) - ${student["비고"]}</div>
                <div class="class-info">
                    출신 반: <span class="origin-class" style="background-color:${getClassColor(student["전 학년 반"])};">${student["전 학년 반"]}반</span> 
                </div>
            `;
            studentDiv.addEventListener('dragstart', handleDragStart);
            classBox.appendChild(studentDiv);
        });

        container.appendChild(classBox);
    }
}

function getClassColor(className) {
    const classColors = {
        "1": "#FFD700", // Yellow
        "2": "#87CEFA", // Sky blue
        "3": "#90EE90", // Light green
        "4": "#FFB6C1", // Light pink
        "5": "#D3D3D3", // Light gray
        "Class_1": "#FFD700",
        "Class_2": "#87CEFA",
        "Class_3": "#90EE90",
        "Class_4": "#FFB6C1",
        "Class_5": "#D3D3D3",
    };
    return classColors[className] || "#FFFFFF"; // Default to white
}

// 성별 색상 적용
function applyGenderColors() {
    const students = document.querySelectorAll('.student');
    students.forEach(student => {
        const gender = student.textContent.match(/\((남|여)\)/); // 성별 추출
        if (gender) {
            if (gender[1] === "남") {
                student.style.backgroundColor = '#ADD8E6'; // 하늘색
                student.style.color = 'black'; // 글씨 색상
            } else if (gender[1] === "여") {
                student.style.backgroundColor = '#FFC0CB'; // 분홍색
                student.style.color = 'black'; // 글씨 색상
            }
        }
    });
}

// 출신 반에 따른 색상 적용
function applyOriginClassColors() {
    const students = document.querySelectorAll('.student');
    const classColors = {}; // 출신 반별 색상 저장
    let colorIndex = 0;

    students.forEach(student => {
        const originClass = student.dataset.originClass; // 출신 반 데이터 (전 학년 반)
        if (!originClass) return; // 출신 반 데이터가 없으면 건너뛰기

        // 출신 반별 색상 생성 (고유 색상 할당)
        if (!classColors[originClass]) {
            classColors[originClass] = `hsl(${(colorIndex * 137.508) % 360}, 70%, 80%)`;
            colorIndex++;
        }

        // 기존 색상 초기화
        student.style.backgroundColor = '';
        student.style.color = '';

        // 새로운 색상 적용
        student.style.backgroundColor = classColors[originClass];
        student.style.color = 'black'; // 텍스트 가독성 확보
    });
}


function resetColors() {
    const students = document.querySelectorAll('.student');
    students.forEach(student => {
        // 모든 인라인 스타일 초기화
        student.style.backgroundColor = '';
        student.style.color = '';

        // 기존 클래스 제거 (출신 반 색상 관련 클래스 등)
        student.classList.remove(...Array.from(student.classList).filter(cls => cls.startsWith('origin-class-')));
    });
}


function renderSummaryLine(label, currentValue, totalStudents, numClasses) {
    // 기준값 계산
    const baseline = Math.floor(totalStudents / numClasses);
    const tolerance = 1; // ±1 범위 허용
    const difference = currentValue - baseline;

    // 색상 결정
    let color = "black";
    if (difference > tolerance) color = "blue"; // 기준값보다 크게 초과
    else if (difference < -tolerance) color = "red"; // 기준값보다 크게 미달

    // UI 생성
    return `
        <p style="color: ${color};">
            ${label}: ${currentValue} (기준: ${baseline} ±${tolerance}, 변화: ${difference >= 0 ? "+" : ""}${difference})
        </p>
    `;
}



// Allow drop
function allowDrop(event) {
    event.preventDefault();
}

// Handle drag start
function handleDragStart(event) {
    const studentId = event.target.getAttribute('data-student-id');
    const className = event.target.parentNode.getAttribute('data-class-name');
    event.dataTransfer.setData('text/plain', JSON.stringify({ studentId, className }));
}

// Handle drop
function handleDrop(event) {
    event.preventDefault();
    const data = JSON.parse(event.dataTransfer.getData('text/plain'));
    const sourceClass = data.className;
    const studentId = parseInt(data.studentId);
    const targetClass = event.currentTarget.getAttribute('data-class-name');


    // Move student
    const student = classesData[sourceClass].students.splice(studentId, 1)[0];
    classesData[targetClass].students.push(student);

    // Update summaries
    updateSummary(sourceClass);
    updateSummary(targetClass);

    // Re-render
    renderClasses();

    const splitResults = checkSplitConditions(currentSplitData, classesData);
    renderSplitConditions(splitResults); // Update the split condition UI

    // Send updated data to server
    saveChanges();
}

// Update class summary
function updateSummary(className) {
    const classData = classesData[className];
    const maleCount = classData.students.filter(student => student["성별"] === "남").length;
    const femaleCount = classData.students.filter(student => student["성별"] === "여").length;
    const totalScore = classData.students.reduce((sum, student) => sum + calculateStudentScore(student), 0);

    classData.summary["총 학생 수"] = classData.students.length;
    classData.summary["남학생 수"] = maleCount;
    classData.summary["여학생 수"] = femaleCount;
    classData.summary["학급 전체 점수"] = totalScore;
}

// Calculate student score
function calculateStudentScore(student) {
    const weights = { "성적 등급 (A/B/C/D)": 1, "생활지도 어려움 등급 (A/B/C/D)": 2 };
    let score = 0;
    for (const [key, weight] of Object.entries(weights)) {
        if (student[key] && "ABCD".includes(student[key])) {
            score += weight * (student[key].charCodeAt(0) - "A".charCodeAt(0) + 1);
        }
    }
    return score;
}

function calculateTotalStudents(classesData) {
    let totalMaleStudents = 0;
    let totalFemaleStudents = 0;

    for (const classData of Object.values(classesData)) {
        totalMaleStudents += classData.summary["남학생 수"] || 0;
        totalFemaleStudents += classData.summary["여학생 수"] || 0;
    }

    return { totalMaleStudents, totalFemaleStudents };
}

function calculateNumClasses(classesData) {
    return Object.keys(classesData).length;
}

function checkSplitConditions(splitData, classesData) {
    const results = [];

    splitData.forEach(group => {
        const groupDetails = group.map(id => {
            const classNumber = Math.floor(id / 100);
            const studentNumber = id % 100;

            // 학생 데이터 찾기
            for (const [className, classData] of Object.entries(classesData)) {
                const student = classData.students.find(
                    s => s["반"] === classNumber && s["번호"] === studentNumber
                );
                if (student) {
                    return { ...student, currentClass: className };
                }
            }

            return null; // 학생을 찾을 수 없을 경우
        });

        // 같은 반에 있는지 확인
        const currentClasses = new Set(groupDetails.map(s => s?.currentClass).filter(Boolean));
        const isSeparated = currentClasses.size === groupDetails.length;

        results.push({ groupDetails, isSeparated });
    });

    return results;
}

function renderSplitConditions(splitResults) {
    const container = document.getElementById('split-conditions');
    container.innerHTML = ''; // 기존 내용 초기화

    splitResults.forEach((result, index) => {
        const groupDiv = document.createElement('div');
        groupDiv.classList.add('split-group');
        groupDiv.style.color = result.isSeparated ? 'black' : 'red'; // 분리 실패 시 빨간색

        groupDiv.innerHTML = `
            <h4>Group ${index + 1} (${result.isSeparated ? 'Separated' : 'Not Separated'})</h4>
            <ul>
                ${result.groupDetails
                    .map(
                        student =>
                            `<li>${student ? `${student["학생 이름"]} (${student["반"]}반 ${student["번호"]}번 -> ${student.currentClass})` : 'Unknown Student'}</li>`
                    )
                    .join('')}
            </ul>
        `;
        container.appendChild(groupDiv);
    });
}


async function saveChanges() {
    console.log("Sending data to server:", classesData); // 디버깅용 데이터 출력
    const response = await fetch('/update_classes', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(classesData),
    });
    const result = await response.json();
    if (response.ok) {
        console.log("Server response:", result);
    } else {
        alert(`Error: ${result.error}`);
    }
}

async function saveToExcel() {
    const response = await fetch('/download_excel', {
        method: 'GET', // POST 대신 GET으로 변경
    });
    if (response.ok) {
        // ZIP 파일 다운로드 처리
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'Class_Data.zip'; // 다운로드할 파일 이름
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
    } else {
        const result = await response.json();
        alert(`Error: ${result.error}`); // 오류 메시지
    }
}



// Fetch initial data
async function fetchClasses() {
    const response = await fetch('/get_classes');
    classesData = await response.json();

    numClasses = calculateNumClasses(classesData)-1;    

    // 원본 요약 정보 저장
    for (const [className, classData] of Object.entries(classesData)) {
        classData.originalSummary = { ...classData.summary }; // 초기 요약 정보를 복사
    }

    renderClasses();
}

async function fetchSplitData() {
    try {
        const response = await fetch('/get_split_data');
        if (!response.ok) throw new Error("Failed to fetch split data.");
        return await response.json();
    } catch (error) {
        console.error("Error loading split data:", error);
        return [];
    }
}

async function loadSplitConditions() {
    try {
        currentSplitData = await fetchSplitData(); // 전역 변수에 저장
        const response = await fetch('/get_classes');
        if (!response.ok) throw new Error("Failed to fetch class data.");

        const classesData = await response.json();
        const splitResults = checkSplitConditions(currentSplitData, classesData);

        renderSplitConditions(splitResults);
    } catch (error) {
        console.error("Error loading split conditions:", error);
    }
}

// Modal toggle logic
function toggleModal() {
    const modal = document.getElementById('split-conditions-modal');
    modal.style.display = modal.style.display === "block" ? "none" : "block";
}

// Close button event listener
document.querySelector('.close-button').addEventListener('click', () => {
    document.getElementById('split-conditions-modal').style.display = "none";
});

// Escape key to close modal
window.addEventListener('keydown', (event) => {
    if (event.key === "Escape") {
        document.getElementById('split-conditions-modal').style.display = "none";
    }
});

// Call this function to update and show the modal
function showSplitConditions(splitResults) {
    renderSplitConditions(splitResults); // Update UI
    toggleModal(); // Show modal
}


// 페이지 로드 시 실행
loadSplitConditions();


// Initialize
fetchClasses();

document.getElementById('save-excel-button').addEventListener('click', saveToExcel);

document.getElementById('gender-color-button').addEventListener('click', applyGenderColors);
document.getElementById('origin-class-color-button').addEventListener('click', applyOriginClassColors);
document.getElementById('reset-color-button').addEventListener('click', resetColors);

document.getElementById('show-color-legend').addEventListener('click', () => {
    const legend = document.getElementById('color-legend');
    legend.style.display = legend.style.display === 'none' ? 'block' : 'none';
});


document.getElementById('show-split-conditions').addEventListener('click', () => {
    toggleModal(); // Modal 열기
});
