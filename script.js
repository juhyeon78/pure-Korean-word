document.addEventListener('DOMContentLoaded', () => {
    // !!! 중요: 배포된 자신의 Google Apps Script 웹 앱 URL로 변경하세요.
    const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbw0Y0BAVUAv7OZTICBKR7LAonfd1ToD8J-kPGP_8e7IzhHoGGwPCxyNhMqTDCZUr5Vg_A/exec';

    const recordForm = document.getElementById('record-form');
    const recordsContainer = document.getElementById('records-container');
    const exportButton = document.getElementById('export-excel');
    const checkboxError = document.getElementById('checkbox-error');
    let recordsCache = []; // 데이터 캐싱

    // 체크박스 개수 제한 (정확히 2개만 선택)
    const preferenceCheckboxes = document.querySelectorAll('input[name="preference"]');
    preferenceCheckboxes.forEach(checkbox => {
        checkbox.addEventListener('change', () => {
            const checkedCount = document.querySelectorAll('input[name="preference"]:checked').length;
            
            // 2개 초과 선택 방지
            if (checkedCount > 2) {
                checkbox.checked = false;
                checkboxError.style.display = 'block';
                setTimeout(() => {
                    checkboxError.style.display = 'none';
                }, 2000);
            } else {
                checkboxError.style.display = 'none';
            }
        });
    });

    // 데이터 로드 및 화면 업데이트
    const loadRecords = async () => {
        try {
            recordsContainer.innerHTML = '<p style="text-align: center; padding: 30px; color: #5DADE2;">🌿 데이터를 불러오는 중...</p>';
            
            const response = await fetch(WEB_APP_URL, { method: 'GET', redirect: 'follow' });
            if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
            recordsCache = await response.json();

            // 서버에서 받은 데이터가 배열인지 확인
            if (!Array.isArray(recordsCache)) {
                console.error("Error data received from Google Apps Script:", recordsCache);
                throw new Error('Google Apps Script에서 에러가 발생했습니다. 개발자 도구(F12)의 Console 탭에서 상세 정보를 확인하세요.');
            }
            
            // 최신순으로 정렬 (Timestamp 기준)
            recordsCache.sort((a, b) => new Date(b.Timestamp) - new Date(a.Timestamp));
            
            recordsContainer.innerHTML = ''; // 로딩 메시지 제거
            
            if (recordsCache.length === 0) {
                recordsContainer.innerHTML = '<p style="text-align: center; padding: 30px; color: #76D7C4;">🍃 아직 작성된 의견이 없습니다. 첫 번째 의견을 남겨주세요!</p>';
            } else {
                recordsCache.forEach(addRecordToDOM);
            }

        } catch (error) {
            console.error('Error loading records:', error);
            recordsContainer.innerHTML = `<p style="color: #E74C3C; text-align: center; padding: 30px;">⚠️ 데이터를 불러오는 데 실패했습니다.<br>Google Apps Script 설정을 확인해주세요.<br><small>${error.message}</small></p>`;
        }
    };

    // DOM에 기록 목록 행 추가
    const addRecordToDOM = (record) => {
        const row = document.createElement('div');
        row.classList.add('record-row');

        // 날짜 포맷팅
        const date = new Date(record.Timestamp);
        const formattedDate = `${date.getFullYear()}.${String(date.getMonth() + 1).padStart(2, '0')}.${String(date.getDate()).padStart(2, '0')}`;

        row.innerHTML = `
            <div class="record-feeling">${record.Feeling}</div>
            <div class="record-know">${record.Know}</div>
            <div class="record-word">${record.Word || '-'}</div>
            <div class="record-curious" title="${record.Curious}">${record.Curious}</div>
            <div class="record-preference" title="${record.Preference}">${record.Preference}</div>
            <div class="record-date">${formattedDate}</div>
        `;
        recordsContainer.appendChild(row);
    };

    // 폼 제출 이벤트 처리
    recordForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        
        // 체크박스 검증 (정확히 2개 선택)
        const checkedPreferences = document.querySelectorAll('input[name="preference"]:checked');
        if (checkedPreferences.length !== 2) {
            checkboxError.style.display = 'block';
            checkboxError.textContent = '⚠️ 정확히 2개를 선택해주세요!';
            setTimeout(() => {
                checkboxError.style.display = 'none';
            }, 3000);
            return;
        }

        const submitButton = e.target.querySelector('button[type="submit"]');
        submitButton.disabled = true;
        submitButton.textContent = '🌱 소중한 의견 담는 중...';

        const formData = new FormData(recordForm);
        
        // 체크박스 값들을 배열로 수집
        const preferences = [];
        checkedPreferences.forEach(checkbox => {
            preferences.push(checkbox.value);
        });

        const data = {
            feeling: formData.get('feeling'),
            know: formData.get('know'),
            word: formData.get('word') || '',
            curious: formData.get('curious'),
            preference: preferences.join(', ') // 배열을 문자열로 변환
        };

        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                mode: 'no-cors', // Apps Script는 no-cors 모드 필요
                cache: 'no-cache',
                redirect: 'follow',
                body: JSON.stringify(data)
            });

            // no-cors 모드에서는 응답을 직접 읽을 수 없으므로, 성공적으로 전송되었다고 가정
            alert('✨ 소중한 의견이 성공적으로 저장되었습니다!');
            recordForm.reset();
            
            // 약간의 지연 후 데이터 다시 불러오기 (서버 처리 시간 고려)
            setTimeout(() => {
                loadRecords();
            }, 1000);

        } catch (error) {
            console.error('Error submitting record:', error);
            alert('❌ 의견 저장에 실패했습니다. 인터넷 연결을 확인하세요.');
        } finally {
            submitButton.disabled = false;
            submitButton.textContent = '소중한 의견 담기 🌱';
        }
    });

    // 엑셀 내보내기 이벤트 처리
    exportButton.addEventListener('click', () => {
        if (recordsCache.length === 0) {
            alert('📭 내보낼 데이터가 없습니다.');
            return;
        }

        try {
            // 엑셀용 데이터 정리
            const excelData = recordsCache.map(record => ({
                '작성일시': new Date(record.Timestamp).toLocaleString('ko-KR'),
                '순우리말 느낌': record.Feeling,
                '아는 단어 여부': record.Know,
                '아는 단어': record.Word || '-',
                '궁금한 의미': record.Curious,
                '원하는 기능': record.Preference
            }));

            // 데이터 시트 생성
            const worksheet = XLSX.utils.json_to_sheet(excelData);
            
            // 컬럼 너비 설정
            worksheet['!cols'] = [
                { wch: 20 }, // 작성일시
                { wch: 15 }, // 순우리말 느낌
                { wch: 15 }, // 아는 단어 여부
                { wch: 20 }, // 아는 단어
                { wch: 30 }, // 궁금한 의미
                { wch: 50 }  // 원하는 기능
            ];

            // 새 워크북 생성
            const workbook = XLSX.utils.book_new();
            // 워크북에 데이터 시트 추가
            XLSX.utils.book_append_sheet(workbook, worksheet, "순우리말사전_설문");

            // 엑셀 파일 내보내기
            const fileName = `순우리말사전_설문결과_${new Date().toISOString().split('T')[0]}.xlsx`;
            XLSX.writeFile(workbook, fileName);
            
            alert('✅ 엑셀 파일이 다운로드되었습니다!');
        } catch (error) {
            console.error('Excel export error:', error);
            alert('❌ 엑셀 내보내기에 실패했습니다.');
        }
    });

    // 초기 데이터 로드
    loadRecords();
});
