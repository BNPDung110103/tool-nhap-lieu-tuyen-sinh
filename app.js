const langSelect = document.getElementById('lang-select');
const dataTableBody = document.getElementById('table-body');
const recordCountEl = document.getElementById('record-count');
const addRecordBtn = document.getElementById('add-record-btn');
const clearFormBtn = document.getElementById('clear-form-btn');
const smartCaptureBtn = document.getElementById('smart-capture-btn');
const exportCsvBtn = document.getElementById('export-csv-btn');
const globalStatus = document.getElementById('global-recording-status');
const stopGlobalMicBtn = document.getElementById('stop-global-mic');
const indicatorText = document.getElementById('listening-indicator-text');
const majorMatchStatus = document.getElementById('major-match-status');

// Form inputs Map
const inputsMap = {
    'f-name': document.getElementById('f-name'),
    'f-dob': document.getElementById('f-dob'),
    'f-phone': document.getElementById('f-phone'),
    'f-address': document.getElementById('f-address'),
    'f-major': document.getElementById('f-major'),
    'f-tuition': document.getElementById('f-tuition')
};

// 53 Majors and their 1-credit tuition mapping
const tuitionMap = {
    "tiếp viên hàng không": 590000, "dịch vụ thương mại hàng không": 590000,
    "điện công nghiệp": 525000, "điện tử công nghiệp": 525000,
    "công nghệ kỹ thuật điện, điện tử": 525000, "công nghệ kỹ thuật điện điện tử": 525000,
    "kỹ thuật máy lạnh và điều hoà không khí": 535000, "vi mạch bán dẫn": 535000,
    "tiếng anh thương mại": 525000, "phiên dịch tiếng anh thương mại": 525000,
    "phiên dịch tiếng anh du lịch": 525000, "phiên dịch tiếng đức kinh tế, thương mại": 525000,
    "phiên dịch tiếng đức kinh tế thương mại": 525000, "tiếng hàn quốc": 525000, "tiếng trung quốc": 525000, "tiếng nhật": 525000,
    "y sĩ đa khoa": 565000, "y học cổ truyền": 565000, "dược": 500000, "điều dưỡng": 500000,
    "hộ sinh": 565000, "kỹ thuật hình ảnh y học": 565000, "kỹ thuật phục hình răng": 565000, "kỹ thuật phục hồi chức năng": 565000,
    "dịch vụ chăm sóc gia đình": 535000, "công nghệ thông tin (ưdpm)": 525000, "công nghệ thông tin ưdpm": 525000,
    "công nghệ thông tin": 525000, "thiết kế đồ họa": 525000, "marketing thương mại": 510000,
    "thương mại điện tử": 535000, "công nghệ kỹ thuật ô tô": 535000, "hàn": 525000,
    "chăm sóc sắc đẹp": 535000, "kỹ thuật chế biến món ăn": 535000, "thú y": 535000,
    "tiếng việt và văn hóa việt nam": 535000, "báo chí": 535000, "quan hệ công chúng": 525000,
    "truyền thông đa phương tiện": 525000, "pháp luật": 535000, "luật dvpl doanh nghiệp": 535000,
    "luật dvpl về đất đai": 535000, "luật dvpl về tố tụng": 535000, "kế toán doanh nghiệp": 535000,
    "quản trị doanh nghiệp vừa và nhỏ": 535000, "logistic": 535000, "logistics": 535000,
    "quản trị kinh doanh": 535000, "hướng dẫn du lịch": 535000, "quản trị khách sạn": 535000,
    "quản trị nhà hàng": 535000, "thiết kế thời trang": 525000, "may thời trang": 525000,
    "văn thư hành chính": 535000, "quản trị văn phòng": 510000, "văn thư - lưu trữ": 535000, "văn thư lưu trữ": 535000
};

// Remove diacritics for better matching functionality
function removeAscii(str) {
    return str.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd');
}

// Map stripped key to standard keys
const normalizedTuitionMap = {};
for (const [key, value] of Object.entries(tuitionMap)) {
    normalizedTuitionMap[removeAscii(key)] = { originalKey: key, tuition: value };
}

// Major Matching Logic
function matchMajor(inputStr) {
    let checkStr = removeAscii(inputStr.trim());
    if(!checkStr) {
        inputsMap['f-tuition'].value = '';
        majorMatchStatus.textContent = '';
        return;
    }
    
    // Exact or partial fuzzy match
    let bestMatch = null;
    if (normalizedTuitionMap[checkStr]) {
        bestMatch = normalizedTuitionMap[checkStr];
    } else {
        // Simple fallback: check if one includes the other
        for (const [normKey, data] of Object.entries(normalizedTuitionMap)) {
            if (normKey.includes(checkStr) || checkStr.includes(normKey)) {
                bestMatch = data;
                break;
            }
        }
    }

    if (bestMatch) {
         // Auto capitalize original key properly for Vietnamese
         let capsKey = capitalizeWords(bestMatch.originalKey);
         inputsMap['f-major'].value = capsKey; // optionally replace their text with standardized
         inputsMap['f-tuition'].value = bestMatch.tuition;
         majorMatchStatus.textContent = "✔ Tìm thấy ngành khớp";
         majorMatchStatus.className = 'match-status success';
    } else {
         inputsMap['f-tuition'].value = '';
         majorMatchStatus.textContent = "⚠ Không thấy ngành phù hợp";
         majorMatchStatus.className = 'match-status error';
    }
}

inputsMap['f-major'].addEventListener('input', (e) => matchMajor(e.target.value));

function capitalizeWords(str) {
    return str.replace(/\S+/g, w => w.charAt(0).toUpperCase() + w.substring(1).toLowerCase());
}

function parseSmartCapture(text) {
    let lowerText = text.toLowerCase();
    
    // Define extraction markers
    const markers = [
        { key: 'name', labels: ['họ và tên', 'họ tên', 'tên là', 'tên'] },
        { key: 'dob', labels: ['ngày sinh', 'sinh ngày', 'sinh mùng', 'sinh'] },
        { key: 'phone', labels: ['số điện thoại', 'sđt', 'điện thoại'] },
        { key: 'address', labels: ['địa chỉ', 'thường trú', 'ở', 'địa chỉ ở'] },
        { key: 'major', labels: ['ngành học', 'ngành', 'đăng ký ngành'] }
    ];

    let foundMarkers = [];
    
    markers.forEach(m => {
        for(let label of m.labels) {
            let idx = lowerText.indexOf(label);
            if(idx !== -1) {
                foundMarkers.push({
                    key: m.key,
                    label: label,
                    index: idx,
                    length: label.length
                });
                break; 
            }
        }
    });

    foundMarkers.sort((a, b) => a.index - b.index);

    // Extract blocks between keywords
    for(let i = 0; i < foundMarkers.length; i++) {
        let current = foundMarkers[i];
        let next = foundMarkers[i+1];
        
        let start = current.index + current.length;
        let end = next ? next.index : text.length;
        
        let value = text.substring(start, end).trim();
        
        // Cleanup prefix artifacts like 'là ' or ': '
        if (value.startsWith('là ')) value = value.substring(3).trim();
        if (value.startsWith(': ')) value = value.substring(2).trim();

        // Assign
        if(current.key === 'name') inputsMap['f-name'].value = capitalizeWords(value);
        if(current.key === 'dob') {
            inputsMap['f-dob'].value = value.replace(/tháng/g, '/').replace(/năm/g, '/').replace(/mùng/g,'').replace(/ /g,'');
        }
        if(current.key === 'phone') inputsMap['f-phone'].value = value.replace(/[^0-9]/g, '');
        if(current.key === 'address') inputsMap['f-address'].value = capitalizeWords(value);
        if(current.key === 'major') {
            inputsMap['f-major'].value = value;
            matchMajor(value);
        }
    }
}

/// Speech Recognition System ///
let recognition;
let isRecording = false;
let currentTargetId = null; 

function initSpeechRecognition() {
    window.SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!window.SpeechRecognition) {
        alert('Trình duyệt không hỗ trợ Web Speech API. Vui lòng dùng Chrome/Edge.');
        return;
    }
    recognition = new window.SpeechRecognition();
    recognition.continuous = true;
    recognition.interimResults = true;
    recognition.lang = langSelect.value;
    
    recognition.onstart = () => {
        isRecording = true;
        globalStatus.classList.remove('hidden');
    };

    recognition.onresult = (event) => {
        let interimTranscript = '';
        let finalTranscript = '';

        for (let i = event.resultIndex; i < event.results.length; ++i) {
            if (event.results[i].isFinal) {
                finalTranscript += event.results[i][0].transcript;
            } else {
                interimTranscript += event.results[i][0].transcript;
            }
        }

        if (finalTranscript && currentTargetId) {
            let rawStr = finalTranscript.trim();

            if (currentTargetId === 'smart-capture') {
                parseSmartCapture(rawStr);
                return;
            }

            let el = inputsMap[currentTargetId];
            if (currentTargetId === 'f-dob') {
                 // Try to format dates, "mùng 1 tháng 1 năm 2000" -> logic is complex, keep exact text for safety or basic replace
                 rawStr = rawStr.replace(/tháng/g, '/').replace(/năm/g, '/').replace(/mùng/g,'').replace(/ /g,'');
            } 
            if (currentTargetId === 'f-phone') {
                 rawStr = rawStr.replace(/[^0-9]/g, '');
            }

            if (el.value && !el.value.endsWith(' ')) {
                el.value += ' ' + rawStr;
            } else {
                el.value += rawStr;
            }
            
            // Clean up values
            el.value = el.value.trim();

            if (currentTargetId === 'f-major') {
                matchMajor(el.value);
            }
        }
        
        let labelName = document.querySelector(`label[for="${currentTargetId}"]`)?.innerText || 'ô nhập liệu';
        indicatorText.innerText = interimTranscript ? `Nghe: ${interimTranscript}` : `Đang thu âm cho [${labelName}]...`;
    };

    recognition.onerror = (event) => { stopGlobalMic(); };
    recognition.onend = () => { if (isRecording) { try { recognition.start(); } catch(e){ stopGlobalMic(); } } else { stopGlobalMic(); } };
}

function startRecordingFor(targetId) {
    if(isRecording && currentTargetId === targetId) { stopGlobalMic(); return; }
    
    if(!recognition) initSpeechRecognition();
    if(isRecording) recognition.stop(); // Stop current one
    
    currentTargetId = targetId;
    recognition.lang = langSelect.value;
    
    // UI logic for mic buttons
    document.querySelectorAll('.mic-btn-small').forEach(btn => btn.classList.remove('listening'));
    
    let targetBtn = document.querySelector(`button[data-target="${targetId}"]`);
    if(targetBtn) {
        targetBtn.classList.add('listening');
    }

    setTimeout(() => { try { isRecording = true; recognition.start(); } catch(e){} }, 100);
}

function stopGlobalMic() {
    isRecording = false;
    currentTargetId = null;
    if(recognition) recognition.stop();
    globalStatus.classList.add('hidden');
    document.querySelectorAll('.mic-btn-small').forEach(btn => btn.classList.remove('listening'));
}

document.querySelectorAll('.mic-btn-small').forEach(btn => {
    btn.addEventListener('click', (e) => {
        let target = e.currentTarget.getAttribute('data-target');
        startRecordingFor(target);
    });
});

if(smartCaptureBtn) {
    smartCaptureBtn.addEventListener('click', () => {
        startRecordingFor('smart-capture');
    });
}

stopGlobalMicBtn.addEventListener('click', stopGlobalMic);
langSelect.addEventListener('change', () => { if(recognition) recognition.lang = langSelect.value; });

/// Table Functionality ///
let recordCount = 0;

function clearForm() {
    Object.values(inputsMap).forEach(el => el.value = '');
    majorMatchStatus.textContent = '';
}

clearFormBtn.addEventListener('click', clearForm);

addRecordBtn.addEventListener('click', () => {
    let name = inputsMap['f-name'].value;
    let dob = inputsMap['f-dob'].value;
    let phone = inputsMap['f-phone'].value;
    let address = inputsMap['f-address'].value;
    let major = inputsMap['f-major'].value;
    let tuition = inputsMap['f-tuition'].value;

    if(!name && !phone) {
        alert("Vui lòng nhập điền ít nhất Tên hoặc Số điện thoại để thêm."); return;
    }

    recordCount++;
    const row = document.createElement('tr');
    row.innerHTML = `
        <td>${recordCount}</td>
        <td>${name}</td>
        <td>${dob}</td>
        <td>${phone}</td>
        <td>${address}</td>
        <td>${major}</td>
        <td>${tuition}</td>
        <td><button class="delete-row-btn" title="Xóa dòng này"><i class="fa-solid fa-xmark"></i></button></td>
    `;
    
    row.querySelector('.delete-row-btn').addEventListener('click', () => {
        row.remove();
        // Option to re-index here but keeping original index is fine too
        updateRecordCount();
    });

    dataTableBody.appendChild(row);
    updateRecordCount();
    clearForm();
    
    // Auto focus back on first element for speed
    inputsMap['f-name'].focus();
});

function updateRecordCount() {
    const rows = dataTableBody.querySelectorAll('tr');
    recordCountEl.textContent = rows.length;
    recordCount = rows.length; // Đồng bộ biến đếm để tạo dòng mới chính xác
    
    // Đánh lại Số thứ tự (STT) cho tất cả các dòng còn lại
    rows.forEach((row, index) => {
        let firstCell = row.querySelector('td:first-child');
        if (firstCell) {
            firstCell.textContent = index + 1;
        }
    });
}

/// Excel (XLSX) Export Functionality ///
exportCsvBtn.addEventListener('click', () => {
    // Requires SheetJS (xlsx.js) included in HTML
    let rows = dataTableBody.querySelectorAll('tr');
    if(rows.length === 0) { alert('Không có dữ liệu để xuất!'); return; }
    
    let data = [
        ["STT", "Họ và Tên", "Ngày Sinh", "Số Điện Thoại", "Địa Chỉ", "Ngành Học", "Học Phí"]
    ];
    
    rows.forEach(row => {
        let cols = row.querySelectorAll('td');
        let rowData = Array.from(cols).slice(0, 7).map(col => col.innerText);
        data.push(rowData);
    });
    
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "DanhSach");
    
    XLSX.writeFile(wb, `Data_TuyenSinh_${new Date().toISOString().slice(0,10)}.xlsx`);
});

// Initialize Speech Object empty
initSpeechRecognition();
