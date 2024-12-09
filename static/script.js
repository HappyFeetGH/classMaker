let classesData = {};

// Fetch initial data
async function fetchClasses() {
    const response = await fetch('/get_classes');
    classesData = await response.json();
    renderClasses();
}

// Render classes
function renderClasses() {
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
        classHeader.textContent = className;
        classBox.appendChild(classHeader);

        // 학급 요약 정보 추가
        const summary = document.createElement('div');
        summary.classList.add('class-summary');
        summary.innerHTML = `
            ${renderSummaryLine("총 학생 수", classData.summary["총 학생 수"], classData.originalSummary["총 학생 수"])}
            ${renderSummaryLine("남학생 수", classData.summary["남학생 수"], classData.originalSummary["남학생 수"])}
            ${renderSummaryLine("여학생 수", classData.summary["여학생 수"], classData.originalSummary["여학생 수"])}
            <p>학급 전체 점수: ${classData.summary["학급 전체 점수"]}</p>
        `;
        classBox.appendChild(summary);

        // 학생 목록 추가
        classData.students.forEach((student, index) => {
            const studentDiv = document.createElement('div');
            studentDiv.classList.add('student');
            studentDiv.setAttribute('data-student-id', index); // 학생 인덱스 저장
            studentDiv.setAttribute('draggable', true);
            studentDiv.textContent = `${student["학생 이름"]} (${student["성별"]}) - ${student["비고"]}`;
            studentDiv.addEventListener('dragstart', handleDragStart);
            classBox.appendChild(studentDiv);
        });

        container.appendChild(classBox);
    }
}

// Render a single summary line with comparison
function renderSummaryLine(label, currentValue, originalValue) {
    const difference = currentValue - originalValue;
    let color = "black";
    if (difference > 0) color = "blue"; // 값이 증가한 경우
    else if (difference < 0) color = "red"; // 값이 감소한 경우

    return `
        <p style="color: ${color};">
            ${label}: ${currentValue} (${difference >= 0 ? "+" : ""}${difference})
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

// Fetch initial data
async function fetchClasses() {
    const response = await fetch('/get_classes');
    classesData = await response.json();

    // 원본 요약 정보 저장
    for (const [className, classData] of Object.entries(classesData)) {
        classData.originalSummary = { ...classData.summary }; // 초기 요약 정보를 복사
    }

    renderClasses();
}

// Initialize
fetchClasses();
