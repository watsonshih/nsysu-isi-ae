// 全域變數
let currentUser = null;
let userRole = null;
let studentData = null;
let activities = [];
let students = [];
let admissionYears = [];
let editingActivityId = null;
let currentActivityId = null;

// 新增全域變數
let filteredActivities = [];
let filteredStudents = new Set();
let selectedActivities = new Set();
let selectedStudents = new Set();
let currentActivitySort = { field: 'date', direction: 'desc' };
let currentStudentSort = { field: 'admissionYear', direction: 'desc' };

// DOM 元素
const loadingOverlay = document.getElementById('loadingOverlay');
const loginContainer = document.getElementById('loginContainer');
const studentInfoContainer = document.getElementById('studentInfoContainer');
const mainContent = document.getElementById('mainContent');
const adminContent = document.getElementById('adminContent');
const studentContent = document.getElementById('studentContent');

// 載入與隱藏載入畫面
function showLoadingOverlay() {
    loadingOverlay.classList.remove('hidden');
}

function hideLoadingOverlay() {
    loadingOverlay.classList.add('hidden');
}

// Modal 顯示/隱藏函數
function showModal(modalId, boxId) {
    const modal = document.getElementById(modalId);
    const box = document.getElementById(boxId);

    modal.classList.remove('hidden');
    requestAnimationFrame(() => {
        modal.classList.add('opacity-100');
        box.classList.add('opacity-100', 'scale-100');
        box.classList.remove('scale-95');
    });
}

function hideModal(modalId, boxId) {
    const modal = document.getElementById(modalId);
    const box = document.getElementById(boxId);

    modal.classList.remove('opacity-100');
    box.classList.remove('opacity-100', 'scale-100');
    box.classList.add('scale-95');

    setTimeout(() => {
        modal.classList.add('hidden');
    }, 300);
}

// 確認對話框
function showConfirmModal(message) {
    return new Promise((resolve) => {
        const modal = document.getElementById('confirmModal');
        const box = document.getElementById('confirmBox');
        const msg = document.getElementById('confirmMessage');
        const confirmBtn = document.getElementById('confirmBtn');
        const cancelBtn = document.getElementById('cancelBtn');

        msg.textContent = message;
        showModal('confirmModal', 'confirmBox');

        const cleanup = () => {
            hideModal('confirmModal', 'confirmBox');
            confirmBtn.removeEventListener('click', onConfirm);
            cancelBtn.removeEventListener('click', onCancel);
        };

        const onConfirm = () => {
            cleanup();
            resolve(true);
        };

        const onCancel = () => {
            cleanup();
            resolve(false);
        };

        confirmBtn.addEventListener('click', onConfirm);
        cancelBtn.addEventListener('click', onCancel);
    });
}

// 提示對話框
function showAlertModal(message) {
    return new Promise((resolve) => {
        const modal = document.getElementById('alertModal');
        const box = document.getElementById('alertBox');
        const msg = document.getElementById('alertMessage');
        const okBtn = document.getElementById('alertOkBtn');

        msg.textContent = message;
        showModal('alertModal', 'alertBox');

        const cleanup = () => {
            hideModal('alertModal', 'alertBox');
            okBtn.removeEventListener('click', onOk);
        };

        const onOk = () => {
            cleanup();
            resolve(true);
        };

        okBtn.addEventListener('click', onOk);
    });
}

// 設定認證監聽器
function setupAuthListeners() {
    const googleLoginBtn = document.getElementById('googleLoginBtn');
    googleLoginBtn.addEventListener('click', () => {
        showLoadingOverlay();
        window.firebase.signInWithPopup(window.firebase.auth, window.firebase.provider)
            .catch(error => {
                console.error('登入錯誤:', error);
                hideLoadingOverlay();
                showAlertModal('登入失敗，請重試');
            });
    });

    const logoutBtn = document.getElementById('logoutBtn');
    logoutBtn.addEventListener('click', () => {
        window.firebase.signOut(window.firebase.auth);
    });

    window.firebase.onAuthStateChanged(window.firebase.auth, async (user) => {
        try {
            if (user) {
                currentUser = user;
                showLoadingOverlay();

                userRole = await checkUserRole(user.email);

                if (userRole === 'new') {
                    hideLoadingOverlay();
                    showStudentInfoInterface();
                } else if (userRole === 'admin') {
                    await showAdminInterface();
                } else {
                    studentData = await checkStudentInfo(user.email);
                    if (studentData) {
                        await showStudentInterface();
                    } else {
                        hideLoadingOverlay();
                        showStudentInfoInterface();
                    }
                }
            } else {
                currentUser = null;
                userRole = null;
                studentData = null;
                hideLoadingOverlay();
                showLoginInterface();
            }
        } catch (error) {
            console.error('認證狀態處理錯誤:', error);
            hideLoadingOverlay();
            // 發生錯誤時，根據當前狀態決定顯示哪個界面
            if (user) {
                showAlertModal('載入資料時發生錯誤，請重新整理頁面');
            } else {
                showLoginInterface();
            }
        }
    });
}

// 顯示登入介面
function showLoginInterface() {
    try {
        loginContainer.style.display = 'flex';
        studentInfoContainer.style.display = 'none';
        mainContent.style.display = 'none';
    } catch (error) {
        console.error('顯示登入介面錯誤:', error);
    }
}

// 顯示學生資訊輸入介面
function showStudentInfoInterface() {
    try {
        loginContainer.style.display = 'none';
        studentInfoContainer.style.display = 'block';
        mainContent.style.display = 'none';
    } catch (error) {
        console.error('顯示學生資訊介面錯誤:', error);
        showLoginInterface();
    }
}

// 顯示管理員介面
async function showAdminInterface() {
    try {
        if (!currentUser) {
            throw new Error('使用者資訊不存在');
        }

        const userProfile = document.getElementById('userProfile');
        userProfile.innerHTML = `
            <img src="${currentUser.photoURL}" alt="${currentUser.displayName}" title="${currentUser.email}">
            <span class="hidden sm:inline-block text-sm">${currentUser.displayName}</span>
        `;

        // 先顯示介面，再載入資料
        loginContainer.style.display = 'none';
        studentInfoContainer.style.display = 'none';
        mainContent.style.display = 'block';
        adminContent.classList.remove('hidden');
        studentContent.classList.add('hidden');

        try {
            // 確保依賴順序正確：先載入年份，再載入其他
            await loadAdmissionYears(); // 包含 updateYearSelects()
            await Promise.all([
                loadActivities(),
                loadStudents()
            ]);

            renderActivitiesList();
            renderStudentsList();
        } catch (dataError) {
            console.error('載入管理員資料錯誤:', dataError);
            showAlertModal('載入部分資料時發生錯誤，部分功能可能無法正常使用');
        }

        hideLoadingOverlay();
    } catch (error) {
        console.error('顯示管理員介面錯誤:', error);
        hideLoadingOverlay();
        showAlertModal('載入管理員介面失敗，請重新登入');
        window.firebase.signOut(window.firebase.auth);
    }
}

// 顯示學生介面
async function showStudentInterface() {
    try {
        if (!currentUser) {
            throw new Error('使用者資訊不存在');
        }

        const userProfile = document.getElementById('userProfile');
        userProfile.innerHTML = `
            <img src="${currentUser.photoURL}" alt="${currentUser.displayName}" title="${currentUser.email}">
            <span class="hidden sm:inline-block text-sm">${currentUser.displayName}</span>
        `;

        // 先顯示介面，再載入資料
        loginContainer.style.display = 'none';
        studentInfoContainer.style.display = 'none';
        mainContent.style.display = 'block';
        adminContent.classList.add('hidden');
        studentContent.classList.remove('hidden');

        try {
            await loadActivities(); // 確保活動已載入
            if (studentData && studentData.id) { // 確保 studentData 和 studentData.id 存在
                renderStudentActivitiesList();
            } else {
                console.warn('學生資料不完整，無法渲染活動列表:', studentData);
                document.getElementById('studentActivitiesList').innerHTML = '<p class="text-gray-400">無法載入您的活動紀錄，請確認學號是否正確綁定。</p>';
                document.getElementById('totalActivities').textContent = '0';
            }
        } catch (dataError) {
            console.error('載入學生資料錯誤:', dataError);
            showAlertModal('載入活動資料時發生錯誤');
        }

        hideLoadingOverlay();
    } catch (error) {
        console.error('顯示學生介面錯誤:', error);
        hideLoadingOverlay();
        showAlertModal('載入學生介面失敗，請重新登入');
        window.firebase.signOut(window.firebase.auth);
    }
}

// 格式化日期函數
function formatDate(dateString) {
    if (!dateString) return 'N/A';
    try {
        const date = new Date(dateString);
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0'); // 月份是從 0 開始的
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    } catch (error) {
        console.error('格式化日期錯誤:', error, '原始日期:', dateString);
        return '日期無效';
    }
}

// 權限檢查
async function checkUserRole(userEmail) {
    try {
        const userEmailKey = userEmail.replace(/\./g, ','); // Firebase keys cannot contain '.'
        const userRef = window.firebase.ref(window.firebase.db, `users/${userEmailKey}`);
        const snapshot = await window.firebase.get(userRef);

        if (snapshot.exists()) {
            const userData = snapshot.val();

            if (userData && userData.role) {
                return userData.role;
            }

            if (userData && userData.studentId) {
                return 'student';
            }
            return 'new';
        }

        return 'new'; // 新使用者或資料庫中無此使用者記錄
    } catch (error) {
        console.error('檢查使用者權限錯誤:', error);
        return 'new';
    }
}

// 檢查學生資訊
async function checkStudentInfo(userEmail) {
    try {
        const userEmailKey = userEmail.replace(/\./g, ',');
        const userRef = window.firebase.ref(window.firebase.db, `users/${userEmailKey}`);
        const snapshot = await window.firebase.get(userRef);

        if (snapshot.exists()) {
            const userData = snapshot.val();
            if (userData.studentId) {
                // 獲取學生詳細資訊
                const studentRef = window.firebase.ref(window.firebase.db, `students/${userData.studentId}`);
                const studentSnapshot = await window.firebase.get(studentRef);
                if (studentSnapshot.exists()) {
                    return studentSnapshot.val();
                }
            }
        }
        return null;
    } catch (error) {
        console.error('檢查學生資訊錯誤:', error);
        return null;
    }
}

// 儲存使用者資訊
async function saveUserInfo(userEmail, userData) {
    try {
        const userEmailKey = userEmail.replace(/\./g, ',');
        const userRef = window.firebase.ref(window.firebase.db, `users/${userEmailKey}`);
        await window.firebase.set(userRef, userData);
        return true;
    } catch (error) {
        console.error('儲存使用者資訊錯誤:', error);
        return false;
    }
}

// 載入入學年資料
async function loadAdmissionYears() {
    try {
        const yearsRef = window.firebase.ref(window.firebase.db, 'admissionYears');
        const snapshot = await window.firebase.get(yearsRef);

        if (snapshot.exists()) {
            const data = snapshot.val();
            admissionYears = Object.keys(data).map(key => ({
                year: key,
                ...data[key]
            })).sort((a, b) => parseInt(b.year) - parseInt(a.year));
        } else {
            admissionYears = [];
        }

        updateYearSelects(); // 移到這裡，確保 admissionYears 已更新
    } catch (error) {
        console.error('載入入學年錯誤:', error);
        admissionYears = []; // 確保有預設值
        updateYearSelects(); // 即使出錯也嘗試更新，避免後續錯誤
        throw error; // 重新拋出錯誤以便上層處理
    }
}

// 載入活動資料
async function loadActivities() {
    try {
        const activitiesRef = window.firebase.ref(window.firebase.db, 'activities');
        const snapshot = await window.firebase.get(activitiesRef);

        if (snapshot.exists()) {
            const data = snapshot.val();
            activities = Object.keys(data).map(key => ({
                id: key,
                ...data[key]
            })).sort((a, b) => new Date(b.date) - new Date(a.date));
        } else {
            activities = [];
        }

        // 更新年份篩選器
        updateActivityYearFilter();
        // 重置篩選結果
        filteredActivities = [...activities];
    } catch (error) {
        console.error('載入活動錯誤:', error);
        activities = [];
        filteredActivities = [];
        throw error;
    }
}

// 載入學生資料
async function loadStudents() {
    try {
        const studentsRef = window.firebase.ref(window.firebase.db, 'students');
        const snapshot = await window.firebase.get(studentsRef);

        if (snapshot.exists()) {
            const data = snapshot.val();
            students = Object.keys(data).map(key => ({
                id: key,
                ...data[key]
            }));
        } else {
            students = [];
        }

        // 重置篩選結果
        filteredStudents = [...students];
    } catch (error) {
        console.error('載入學生錯誤:', error);
        students = [];
        filteredStudents = [];
        throw error;
    }
}

// 更新活動年份篩選器
function updateActivityYearFilter() {
    const yearFilter = document.getElementById('activityYearFilter');
    if (!yearFilter) return;

    // 獲取所有活動的年份
    const years = [...new Set(activities.map(activity => new Date(activity.date).getFullYear()))].sort((a, b) => b - a);

    // 清空選項（保留第一個選項)
    const firstOption = yearFilter.firstElementChild;
    yearFilter.innerHTML = '';
    if (firstOption) {
        yearFilter.appendChild(firstOption);
    }

    // 添加年份選項
    years.forEach(year => {
        const option = document.createElement('option');
        option.value = year;
        option.textContent = `${year}年`;
        yearFilter.appendChild(option);
    });
}

// 更新所有入學年下拉選單
function updateYearSelects() {
    const yearSelects = [
        document.getElementById('yearFilter'),
        document.getElementById('newStudentYear'),
        document.getElementById('uploadStudentYear'),
        document.getElementById('editStudentYear')
    ];

    yearSelects.forEach(selectElement => {
        if (selectElement) {
            const currentValue = selectElement.value; // 保留目前選中的值
            // 清空選項（保留第一個 "請選擇" 或 "所有入學年" 選項)
            const firstOption = selectElement.options[0];
            selectElement.innerHTML = '';
            if (firstOption) {
                selectElement.appendChild(firstOption.cloneNode(true)); // 克隆節點以避免問題
            }

            admissionYears.forEach(yearObj => {
                const option = document.createElement('option');
                option.value = yearObj.year;
                option.textContent = `${yearObj.year} 學年度`;
                selectElement.appendChild(option);
            });

            // 嘗試恢復之前選中的值
            if (currentValue && Array.from(selectElement.options).some(opt => opt.value === currentValue)) {
                selectElement.value = currentValue;
            } else if (selectElement.options.length > 0 && !currentValue) {
                // 如果沒有當前值且有選項，預設選中第一個有效選項（如果不是 "請選擇"）
                if (selectElement.options[0] && selectElement.options[0].value !== "") {
                    // selectElement.value = selectElement.options[0].value; // 避免自動選中
                } else if (selectElement.options.length > 1) {
                    // selectElement.value = selectElement.options[1].value; // 避免自動選中
                }
            }
        }
    });
}

// 設定分頁切換
function setupTabSwitching() {
    const tabButtons = document.querySelectorAll('.tab-button');
    const tabContents = document.querySelectorAll('.tab-content');

    tabButtons.forEach(button => {
        button.addEventListener('click', () => {
            // 移除所有按鈕的 active class
            tabButtons.forEach(btn => btn.classList.remove('active'));
            // 為點擊的按鈕添加 active class
            button.classList.add('active');

            // 隱藏所有內容區域
            tabContents.forEach(content => content.classList.remove('active'));
            // 顯示對應的內容區域
            const tabId = button.dataset.tab;
            const activeContent = document.getElementById(`${tabId}-tab`);
            if (activeContent) {
                activeContent.classList.add('active');
            }
        });
    });
}

// 活動搜尋和篩選
function filterActivities() {
    const searchTerm = document.getElementById('activitySearchInput').value.toLowerCase().trim();
    const statusFilter = document.getElementById('activityStatusFilter').value;
    const yearFilter = document.getElementById('activityYearFilter').value;

    filteredActivities = activities.filter(activity => {
        // 搜尋條件
        const matchesSearch = !searchTerm ||
            activity.name.toLowerCase().includes(searchTerm) ||
            activity.location.toLowerCase().includes(searchTerm) ||
            activity.teacher.toLowerCase().includes(searchTerm);

        // 狀態篩選
        const matchesStatus = !statusFilter ||
            (statusFilter === 'public' && activity.visible) ||
            (statusFilter === 'private' && !activity.visible);

        // 年份篩選
        const matchesYear = !yearFilter ||
            new Date(activity.date).getFullYear().toString() === yearFilter;

        return matchesSearch && matchesStatus && matchesYear;
    });

    // 應用排序
    applyActivitySort();
    renderActivitiesList();
    clearActivitySelection();
}

// 學生搜尋和篩選
function filterStudents() {
    const searchTerm = document.getElementById('studentSearchInput').value.toLowerCase().trim();
    const yearFilter = document.getElementById('yearFilter').value;
    const bindingFilter = document.getElementById('bindingStatusFilter').value;

    filteredStudents = students.filter(student => {
        // 搜尋條件
        const matchesSearch = !searchTerm ||
            student.name.toLowerCase().includes(searchTerm) ||
            student.id.toLowerCase().includes(searchTerm);

        // 入學年篩選
        const matchesYear = !yearFilter || student.admissionYear === yearFilter;

        // 綁定狀態篩選
        const matchesBinding = !bindingFilter ||
            (bindingFilter === 'bound' && student.googleAccount) ||
            (bindingFilter === 'unbound' && !student.googleAccount);

        return matchesSearch && matchesYear && matchesBinding;
    });

    // 應用排序
    applyStudentSort();
    renderStudentsList();
    clearStudentSelection();
}

// 活動排序
function applyActivitySort() {
    filteredActivities.sort((a, b) => {
        const { field, direction } = currentActivitySort;
        let aValue, bValue;

        switch (field) {
            case 'name':
                aValue = a.name;
                bValue = b.name;
                break;
            case 'date':
                aValue = new Date(a.date);
                bValue = new Date(b.date);
                break;
            case 'location':
                aValue = a.location;
                bValue = b.location;
                break;
            case 'teacher':
                aValue = a.teacher;
                bValue = b.teacher;
                break;
            case 'participants':
                aValue = a.participants ? a.participants.length : 0;
                bValue = b.participants ? b.participants.length : 0;
                break;
            case 'visible':
                aValue = a.visible ? 1 : 0;
                bValue = b.visible ? 1 : 0;
                break;
            default:
                return 0;
        }

        if (aValue < bValue) return direction === 'asc' ? -1 : 1;
        if (aValue > bValue) return direction === 'asc' ? 1 : -1;
        return 0;
    });
}

// 學生排序
function applyStudentSort() {
    filteredStudents.sort((a, b) => {
        const { field, direction } = currentStudentSort;
        let aValue, bValue;

        switch (field) {
            case 'name':
                aValue = a.name;
                bValue = b.name;
                break;
            case 'id':
                aValue = a.id;
                bValue = b.id;
                break;
            case 'admissionYear':
                aValue = parseInt(a.admissionYear);
                bValue = parseInt(b.admissionYear);
                break;
            case 'participantCount':
                aValue = getStudentParticipantCount(a.id);
                bValue = getStudentParticipantCount(b.id);
                break;
            case 'googleAccount':
                aValue = a.googleAccount ? 1 : 0;
                bValue = b.googleAccount ? 1 : 0;
                break;
            default:
                return 0;
        }

        if (aValue < bValue) return direction === 'asc' ? -1 : 1;
        if (aValue > bValue) return direction === 'asc' ? 1 : -1;
        return 0;
    });
}

// 渲染活動列表（管理員）- 更新為使用篩選後的資料
function renderActivitiesList() {
    const tableBody = document.getElementById('activitiesTableBody');
    const emptyState = document.getElementById('activitiesEmptyState');

    tableBody.innerHTML = '';

    if (filteredActivities.length === 0) {
        // tableBody.style.display = 'none'; // 改為在CSS中處理或直接移除此行
        emptyState.classList.remove('hidden');
        return;
    }

    // tableBody.style.display = ''; // 改為在CSS中處理或直接移除此行
    emptyState.classList.add('hidden');

    filteredActivities.forEach(activity => {
        const participantCount = activity.participants ? activity.participants.length : 0;
        const isSelected = selectedActivities.has(activity.id);

        const row = document.createElement('tr');
        row.className = 'hover:bg-dark-200 transition-colors';

        row.innerHTML = `
            <td class="px-4 py-3">
                <input type="checkbox" class="activity-checkbox" data-activity-id="${activity.id}" 
                       ${isSelected ? 'checked' : ''} 
                       class="bg-dark-200 border border-dark-300 rounded">
            </td>
            <td class="px-4 py-3 text-gray-200 font-medium">
                <div>
                    <div class="font-semibold">${activity.name}</div>
                    ${activity.notes ? `<div class="text-xs text-gray-400 mt-1">${activity.notes}</div>` : ''}
                </div>
            </td>
            <td class="px-4 py-3 text-gray-300">${formatDate(activity.date)}</td>
            <td class="px-4 py-3 text-gray-300">${activity.location}</td>
            <td class="px-4 py-3 text-gray-300">${activity.teacher}</td>
            <td class="px-4 py-3 text-gray-300">${participantCount} 人</td>
            <td class="px-4 py-3">
                <span class="status-badge ${activity.visible ? 'public' : 'private'}">
                    ${activity.visible ? '公開' : '隱藏'}
                </span>
            </td>
            <td class="px-4 py-3">
                <div class="flex space-x-2">
                    <button class="manage-participants-btn text-purple-400 hover:text-purple-300 p-1 rounded hover:bg-purple-400/10 transition-colors" 
                            data-activity-id="${activity.id}" title="管理參與者">
                        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" 
                                  d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z"></path>
                        </svg>
                    </button>
                    <button class="edit-activity-btn text-blue-400 hover:text-blue-300 p-1 rounded hover:bg-blue-400/10 transition-colors" 
                            data-activity-id="${activity.id}" title="編輯">
                        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" 
                                  d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path>
                        </svg>
                    </button>
                    <button class="delete-activity-btn text-red-400 hover:text-red-300 p-1 rounded hover:bg-red-400/10 transition-colors" 
                            data-activity-id="${activity.id}" title="刪除">
                        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" 
                                  d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path>
                        </svg>
                    </button>
                </div>
            </td>
        `;

        tableBody.appendChild(row);
    });

    // 重新綁定事件監聽器
    bindActivityTableEvents();
    updateActivitySortHeaders();
}

// 渲染學生列表（管理員）- 更新為使用篩選後的資料
function renderStudentsList() {
    const tableBody = document.getElementById('studentsTableBody');
    const emptyState = document.getElementById('studentsEmptyState');

    tableBody.innerHTML = '';

    if (filteredStudents.length === 0) {
        tableBody.style.display = 'none';
        emptyState.classList.remove('hidden');
        return;
    }

    tableBody.style.display = '';
    emptyState.classList.add('hidden');

    filteredStudents.forEach(student => {
        const participantCount = getStudentParticipantCount(student.id);
        const isSelected = selectedStudents.has(student.id);

        const row = document.createElement('tr');
        row.className = 'hover:bg-dark-200 transition-colors';

        row.innerHTML = `
            <td class="px-4 py-3">
                <input type="checkbox" class="student-checkbox" data-student-id="${student.id}" 
                       ${isSelected ? 'checked' : ''} 
                       class="bg-dark-200 border border-dark-300 rounded">
            </td>
            <td class="px-4 py-3 text-gray-200 font-medium">${student.name}</td>
            <td class="px-4 py-3 text-gray-300">${student.id}</td>
            <td class="px-4 py-3 text-gray-300">${student.admissionYear} 學年度</td>
            <td class="px-4 py-3 text-gray-300">${participantCount} 場</td>
            <td class="px-4 py-3 text-gray-300">
                ${student.googleAccount ?
                `<div class="flex items-center gap-2">
                        <span class="text-sm text-gray-400">${student.googleAccount}</span>
                        <span class="inline-block w-2 h-2 bg-green-500 rounded-full" title="已綁定"></span>
                    </div>` :
                `<span class="text-gray-500 text-sm">未綁定</span>`
            }
            </td>
            <td class="px-4 py-3">
                <div class="flex space-x-2">
                    <button class="edit-student-btn text-blue-400 hover:text-blue-300 p-1 rounded hover:bg-blue-400/10 transition-colors" 
                            data-student-id="${student.id}" title="編輯">
                        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" 
                                  d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path>
                        </svg>
                    </button>
                    <button class="delete-student-btn text-red-400 hover:text-red-300 p-1 rounded hover:bg-red-400/10 transition-colors" 
                            data-student-id="${student.id}" title="刪除">
                        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" 
                                  d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path>
                        </svg>
                    </button>
                </div>
            </td>
        `;

        tableBody.appendChild(row);
    });

    // 重新綁定事件監聽器
    bindStudentTableEvents();
    updateStudentSortHeaders();
}

// 綁定活動表格事件
function bindActivityTableEvents() {
    const tableBody = document.getElementById('activitiesTableBody');

    // 移除舊的事件監聽器，避免重複綁定
    const newTableBody = tableBody.cloneNode(false); // 建立一個新的tbody元素但不複製子節點
    while (tableBody.firstChild) {
        newTableBody.appendChild(tableBody.firstChild); // 移動子節點到新的tbody
    }
    tableBody.parentNode.replaceChild(newTableBody, tableBody);


    // 重新綁定事件監聽器到新的tbody
    newTableBody.addEventListener('click', handleActivityTableClick);
    newTableBody.addEventListener('change', handleActivityCheckboxChange);
}

// 綁定學生表格事件
function bindStudentTableEvents() {
    const tableBody = document.getElementById('studentsTableBody');

    // 移除舊的事件監聽器，避免重複綁定
    const newTableBody = tableBody.cloneNode(false);
    while (tableBody.firstChild) {
        newTableBody.appendChild(tableBody.firstChild);
    }
    tableBody.parentNode.replaceChild(newTableBody, tableBody);

    // 重新綁定事件監聽器到新的tbody
    newTableBody.addEventListener('click', handleStudentTableClick);
    newTableBody.addEventListener('change', handleStudentCheckboxChange);
}

// 活動表格點擊處理函數
function handleActivityTableClick(e) {
    const target = e.target.closest('button');
    if (!target) return;

    const activityId = target.dataset.activityId;

    if (target.classList.contains('edit-activity-btn')) {
        editActivity(activityId);
    } else if (target.classList.contains('delete-activity-btn')) {
        deleteActivity(activityId);
    } else if (target.classList.contains('manage-participants-btn')) {
        manageParticipants(activityId);
    }
}

// 學生表格點擊處理函數
function handleStudentTableClick(e) {
    const target = e.target.closest('button');
    if (!target) return;

    const studentId = target.dataset.studentId;

    if (target.classList.contains('edit-student-btn')) {
        editStudent(studentId);
    } else if (target.classList.contains('delete-student-btn')) {
        deleteStudent(studentId);
    }
}

// 活動複選框變更處理
function handleActivityCheckboxChange(e) {
    if (e.target.classList.contains('activity-checkbox')) {
        const activityId = e.target.dataset.activityId;
        if (e.target.checked) {
            selectedActivities.add(activityId);
        } else {
            selectedActivities.delete(activityId);
        }
        updateActivityBulkActions();
        updateSelectAllActivitiesState();
    }
}

// 學生複選框變更處理
function handleStudentCheckboxChange(e) {
    if (e.target.classList.contains('student-checkbox')) {
        const studentId = e.target.dataset.studentId;
        if (e.target.checked) {
            selectedStudents.add(studentId);
        } else {
            selectedStudents.delete(studentId);
        }
        updateStudentBulkActions();
        updateSelectAllStudentsState();
    }
}

// 更新活動批量操作顯示
function updateActivityBulkActions() {
    const bulkActions = document.getElementById('activityBulkActions');
    const countElement = document.getElementById('selectedActivitiesCount');

    if (selectedActivities.size > 0) {
        bulkActions.classList.remove('hidden');
        countElement.textContent = selectedActivities.size;
    } else {
        bulkActions.classList.add('hidden');
    }
}

// 更新學生批量操作顯示
function updateStudentBulkActions() {
    const bulkActions = document.getElementById('studentBulkActions');
    const countElement = document.getElementById('selectedStudentsCount');

    if (selectedStudents.size > 0) {
        bulkActions.classList.remove('hidden');
        countElement.textContent = selectedStudents.size;
    } else {
        bulkActions.classList.add('hidden');
    }
}

// 更新全選活動狀態
function updateSelectAllActivitiesState() {
    const selectAll = document.getElementById('selectAllActivities');
    const displayedActivityIds = filteredActivities.map(a => a.id);
    const selectedDisplayedActivities = displayedActivityIds.filter(id => selectedActivities.has(id));

    if (selectedDisplayedActivities.length === 0) {
        selectAll.indeterminate = false;
        selectAll.checked = false;
    } else if (selectedDisplayedActivities.length === displayedActivityIds.length) {
        selectAll.indeterminate = false;
        selectAll.checked = true;
    } else {
        selectAll.indeterminate = true;
        selectAll.checked = false;
    }
}

// 更新全選學生狀態
function updateSelectAllStudentsState() {
    const selectAll = document.getElementById('selectAllStudents');
    const displayedStudentIds = filteredStudents.map(s => s.id);
    const selectedDisplayedStudents = displayedStudentIds.filter(id => selectedStudents.has(id));

    if (selectedDisplayedStudents.length === 0) {
        selectAll.indeterminate = false;
        selectAll.checked = false;
    } else if (selectedDisplayedStudents.length === displayedStudentIds.length) {
        selectAll.indeterminate = false;
        selectAll.checked = true;
    } else {
        selectAll.indeterminate = true;
        selectAll.checked = false;
    }
}

// 清除活動選擇
function clearActivitySelection() {
    selectedActivities.clear();
    updateActivityBulkActions();
    updateSelectAllActivitiesState();
    document.querySelectorAll('.activity-checkbox').forEach(cb => cb.checked = false);
}

// 清除學生選擇
function clearStudentSelection() {
    selectedStudents.clear();
    updateStudentBulkActions();
    updateSelectAllStudentsState();
    document.querySelectorAll('.student-checkbox').forEach(cb => cb.checked = false);
}

// 批量刪除活動
async function bulkDeleteActivities() {
    if (selectedActivities.size === 0) return;

    const confirmed = await showConfirmModal(`確定要刪除選中的 ${selectedActivities.size} 個活動嗎？此操作無法撤銷。`);
    if (!confirmed) return;

    try {
        showLoadingOverlay();

        const deletePromises = Array.from(selectedActivities).map(activityId => {
            const activityRef = window.firebase.ref(window.firebase.db, `activities/${activityId}`);
            return window.firebase.remove(activityRef);
        });

        await Promise.all(deletePromises);

        // 更新本地資料
        activities = activities.filter(a => !selectedActivities.has(a.id));
        filterActivities();

        clearActivitySelection();
        showAlertModal(`已成功刪除`);
    } catch (error) {
        console.error('批量刪除活動錯誤:', error);
        showAlertModal('批量刪除活動失敗，請重試');
    } finally {
        hideLoadingOverlay();
    }
}

// 批量切換活動顯示狀態
async function bulkToggleActivitiesVisibility() {
    if (selectedActivities.size === 0) return;

    const confirmed = await showConfirmModal(`確定要切換選中的 ${selectedActivities.size} 個活動的顯示狀態嗎？`);
    if (!confirmed) return;

    try {
        showLoadingOverlay();

        const updatePromises = Array.from(selectedActivities).map(activityId => {
            const activity = activities.find(a => a.id === activityId);
            if (activity) {
                const activityRef = window.firebase.ref(window.firebase.db, `activities/${activityId}/visible`);
                return window.firebase.set(activityRef, !activity.visible);
            }
            return Promise.resolve();
        });

        await Promise.all(updatePromises);

        // 更新本地資料
        selectedActivities.forEach(activityId => {
            const activity = activities.find(a => a.id === activityId);
            if (activity) {
                activity.visible = !activity.visible;
            }
        });

        filterActivities();
        clearActivitySelection();
        showAlertModal(`已成功切換 ${selectedActivities.size} 個活動的顯示狀態`);
    } catch (error) {
        console.error('批量切換活動狀態錯誤:', error);
        showAlertModal('批量切換活動狀態失敗，請重試');
    } finally {
        hideLoadingOverlay();
    }
}

// 批量刪除學生
async function bulkDeleteStudents() {
    if (selectedStudents.size === 0) return;

    const confirmed = await showConfirmModal(`確定要刪除選中的 ${selectedStudents.size} 位學生嗎？相關的活動參與紀錄也將被移除。此操作無法撤銷。`);
    if (!confirmed) return;

    try {
        showLoadingOverlay();

        const deletePromises = Array.from(selectedStudents).map(studentId => {
            const studentRef = window.firebase.ref(window.firebase.db, `students/${studentId}`);
            return window.firebase.remove(studentRef);
        });

        await Promise.all(deletePromises);

        // 從所有活動中移除這些學生
        const activityUpdatePromises = activities.map(activity => {
            if (activity.participants) {
                const newParticipants = activity.participants.filter(id => !selectedStudents.has(id));
                if (newParticipants.length !== activity.participants.length) {
                    activity.participants = newParticipants;
                    const activityRef = window.firebase.ref(window.firebase.db, `activities/${activity.id}/participants`);
                    return window.firebase.set(activityRef, newParticipants);
                }
            }
            return Promise.resolve();
        });

        await Promise.all(activityUpdatePromises);

        // 更新本地資料
        students = students.filter(s => !selectedStudents.has(s.id));
        filterStudents();

        clearStudentSelection();
        showAlertModal(`已成功刪除 ${selectedStudents.size} 位學生`);
    } catch (error) {
        console.error('批量刪除學生錯誤:', error);
        showAlertModal('批量刪除學生失敗，請重試');
    } finally {
        hideLoadingOverlay();
    }
}

// 批量解除學生綁定
async function bulkUnbindStudents() {
    if (selectedStudents.size === 0) return;

    const boundStudents = Array.from(selectedStudents).filter(studentId => {
        const student = students.find(s => s.id === studentId);
        return student && student.googleAccount;
    });

    if (boundStudents.length === 0) {
        showAlertModal('選中的學生中沒有已綁定Google帳號的學生');
        return;
    }

    const confirmed = await showConfirmModal(`確定要解除選中的 ${boundStudents.length} 位學生的Google帳號綁定嗎？解除後這些學生將無法登入系統。`);
    if (!confirmed) return;

    try {
        showLoadingOverlay();

        const unbindPromises = boundStudents.map(async studentId => {
            const student = students.find(s => s.id === studentId);
            if (student && student.googleAccount) {
                // 從學生資料中移除googleAccount
                const studentRef = window.firebase.ref(window.firebase.db, `students/${studentId}/googleAccount`);
                await window.firebase.remove(studentRef);

                // 從users表中移除該使用者
                const userEmailKey = student.googleAccount.replace(/\./g, ',');
                const userRef = window.firebase.ref(window.firebase.db, `users/${userEmailKey}`);
                await window.firebase.remove(userRef);

                // 更新本地資料
                delete student.googleAccount;
            }
        });

        await Promise.all(unbindPromises);

        filterStudents();
        clearStudentSelection();
        showAlertModal(`已成功解除 ${boundStudents.length} 位學生的Google帳號綁定`);
    } catch (error) {
        console.error('批量解除綁定錯誤:', error);
        showAlertModal('批量解除綁定失敗，請重試');
    } finally {
        hideLoadingOverlay();
    }
}

// 更新排序標題
function updateActivitySortHeaders() {
    document.querySelectorAll('.activities-table .sortable-header').forEach(header => {
        header.classList.remove('asc', 'desc');
        if (header.dataset.sort === currentActivitySort.field) {
            header.classList.add(currentActivitySort.direction);
        }
    });
}

function updateStudentSortHeaders() {
    document.querySelectorAll('.students-table .sortable-header').forEach(header => {
        header.classList.remove('asc', 'desc');
        if (header.dataset.sort === currentStudentSort.field) {
            header.classList.add(currentStudentSort.direction);
        }
    });
}

// 清除活動篩選
function clearActivityFilters() {
    document.getElementById('activitySearchInput').value = '';
    document.getElementById('activityStatusFilter').value = '';
    document.getElementById('activityYearFilter').value = '';
    filterActivities();
}

// 清除學生篩選
function clearStudentFilters() {
    document.getElementById('studentSearchInput').value = '';
    document.getElementById('yearFilter').value = '';
    document.getElementById('bindingStatusFilter').value = '';
    filterStudents();
}

// 當 Firebase 準備就緒時初始化
document.addEventListener('firebaseReady', () => {
    setupAuthListeners();
    setupTabSwitching(); // 確保此函數已定義並在此處呼叫

    // 活動表單事件
    document.getElementById('newActivityBtn').addEventListener('click', () => {
        showActivityModal(); // 呼叫 showActivityModal
    });

    document.getElementById('cancelActivity').addEventListener('click', () => {
        hideModal('activityModal', 'activityBox');
    });

    document.getElementById('deleteActivity').addEventListener('click', () => {
        if (editingActivityId) {
            deleteActivity(editingActivityId);
            hideModal('activityModal', 'activityBox');
        }
    });

    document.getElementById('activityForm').addEventListener('submit', async (e) => {
        e.preventDefault();

        const formData = new FormData(e.target);
        const activityData = {
            name: formData.get('activityName'),
            date: formData.get('activityDate'),
            location: formData.get('activityLocation'),
            teacher: formData.get('activityTeacher'),
            notes: formData.get('activityNotes'),
            visible: document.getElementById('activityVisible').checked,
            updatedAt: new Date().toISOString()
        };

        try {
            if (editingActivityId) {
                // 更新活動
                const activityRef = window.firebase.ref(window.firebase.db, `activities/${editingActivityId}`);
                await window.firebase.update(activityRef, activityData);

                const index = activities.findIndex(a => a.id === editingActivityId);
                if (index !== -1) {
                    activities[index] = { ...activities[index], ...activityData };
                }

                showAlertModal('活動已更新');
            } else {
                // 新增活動
                activityData.createdAt = new Date().toISOString();
                activityData.participants = [];

                const activitiesRef = window.firebase.ref(window.firebase.db, 'activities');
                const newActivityRef = window.firebase.push(activitiesRef);
                await window.firebase.set(newActivityRef, activityData);

                activities.unshift({
                    id: newActivityRef.key,
                    ...activityData
                });

                showAlertModal('活動已新增');
            }

            hideModal('activityModal', 'activityBox');
            filterActivities();

        } catch (error) {
            console.error('儲存活動錯誤:', error);
            showAlertModal('儲存活動失敗，請重試');
        }
    });

    // 學生資訊表單事件
    document.getElementById('studentInfoForm').addEventListener('submit', async (e) => {
        e.preventDefault();

        const formData = new FormData(e.target);
        const studentId = formData.get('studentId').trim();

        if (!studentId) {
            showAlertModal('請輸入學號');
            return;
        }

        try {
            showLoadingOverlay();

            // 先載入學生資料以查找該學號
            await loadStudents();

            // 查找學號對應的學生資料
            const existingStudent = students.find(s => s.id === studentId);

            if (!existingStudent) {
                showAlertModal('找不到此學號，請確認學號正確或聯繫管理員');
                return;
            }

            // 檢查該學號是否已經綁定其他Google帳號
            if (existingStudent.googleAccount && existingStudent.googleAccount !== currentUser.email) {
                showAlertModal('此學號已綁定其他 Google 帳號，請聯繫管理員');
                return;
            }

            // 如果googleAccount為空，則只更新googleAccount欄位
            if (!existingStudent.googleAccount) {
                const studentRef = window.firebase.ref(window.firebase.db, `students/${studentId}/googleAccount`);
                await window.firebase.set(studentRef, currentUser.email);
            }

            // 建立使用者資料
            const userData = {
                email: currentUser.email,
                displayName: currentUser.displayName,
                photoURL: currentUser.photoURL,
                studentId: studentId,
                role: 'student',
                createdAt: new Date().toISOString()
            };

            await saveUserInfo(currentUser.email, userData);

            showAlertModal('學號綁定成功！');

            // 重新載入資料並顯示學生介面
            studentData = await checkStudentInfo(currentUser.email);
            if (studentData) {
                await showStudentInterface();
            } else {
                throw new Error('綁定後無法載入學生資料');
            }

        } catch (error) {
            console.error('綁定學號錯誤:', error);
            showAlertModal(`綁定失敗：${error.message || '請重試'}`);
        } finally {
            hideLoadingOverlay();
        }
    });

    // 學號綁定頁面登出按鈕
    document.getElementById('logoutFromBinding').addEventListener('click', () => {
        window.firebase.signOut(window.firebase.auth);
    });

    // 學生管理事件
    document.getElementById('newYearBtn').addEventListener('click', () => {
        showModal('yearModal', 'yearBox');
    });

    document.getElementById('cancelYear').addEventListener('click', () => {
        hideModal('yearModal', 'yearBox');
    });

    document.getElementById('yearForm').addEventListener('submit', async (e) => {
        e.preventDefault();
        const year = document.getElementById('newYear').value;

        // 檢查年份是否已存在
        if (admissionYears.find(y => y.year === year)) {
            showAlertModal('此入學年已存在');
            return;
        }

        await addAdmissionYear(year);
    });

    document.getElementById('addStudentBtn').addEventListener('click', () => {
        showModal('studentModal', 'studentBox');
    });

    document.getElementById('cancelStudent').addEventListener('click', () => {
        hideModal('studentModal', 'studentBox');
    });

    document.getElementById('yearFilter').addEventListener('change', () => {
        filterStudents();
    });

    document.getElementById('downloadStudentsBtn').addEventListener('click', () => {
        downloadStudentsSummary();
    });

    // 手動新增學生
    document.getElementById('studentForm').addEventListener('submit', async (e) => {
        e.preventDefault();

        const formData = new FormData(e.target);
        const studentData = {
            id: formData.get('newStudentId'),
            name: formData.get('newStudentName'),
            admissionYear: formData.get('newStudentYear'),
            createdAt: new Date().toISOString()
        };

        await addStudent(studentData);
        e.target.reset();
    });

    // 檔案上傳 - 學生
    const studentFileUpload = document.getElementById('studentFileUpload');
    const studentFileInput = document.getElementById('studentFileInput');
    const uploadStudentFileBtn = document.getElementById('uploadStudentFile');
    const studentFileDisplay = document.getElementById('studentFileDisplay');

    studentFileUpload.addEventListener('click', () => {
        studentFileInput.click();
    });

    studentFileUpload.addEventListener('dragover', (e) => {
        e.preventDefault();
        studentFileUpload.classList.add('dragging');
    });

    studentFileUpload.addEventListener('dragleave', () => {
        studentFileUpload.classList.remove('dragging');
    });

    studentFileUpload.addEventListener('drop', (e) => {
        e.preventDefault();
        studentFileUpload.classList.remove('dragging');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            studentFileInput.files = files;
            displaySelectedFile(studentFileInput, studentFileDisplay, uploadStudentFileBtn);
        }
    });

    studentFileInput.addEventListener('change', (e) => {
        displaySelectedFile(studentFileInput, studentFileDisplay, uploadStudentFileBtn);
    });

    uploadStudentFileBtn.addEventListener('click', async () => {
        const file = studentFileInput.files[0];
        const admissionYear = document.getElementById('uploadStudentYear').value;

        if (!file) {
            showAlertModal('請選擇檔案');
            return;
        }

        if (!admissionYear) {
            showAlertModal('請選擇入學年');
            return;
        }

        try {
            showLoadingOverlay();
            const result = await handleStudentFileUpload(file, admissionYear);

            if (result.errors.length > 0) {
                showAlertModal(`上傳完成，但有 ${result.errors.length} 筆資料有問題：\n${result.errors.join('\n')}`);
            }

            if (result.students.length > 0) {
                // 批量新增學生
                const promises = result.students.map(student => {
                    const studentRef = window.firebase.ref(window.firebase.db, `students/${student.id}`);
                    return window.firebase.set(studentRef, student);
                });

                await Promise.all(promises);
                students.push(...result.students);
                filterStudents(); // 確保篩選和排序被重新應用

                showAlertModal(`成功新增 ${result.students.length} 位學生`);
            }

            clearSelectedFile(studentFileInput, studentFileDisplay, uploadStudentFileBtn);

        } catch (error) {
            console.error('上傳學生檔案錯誤:', error);
            showAlertModal('檔案上傳失敗，請檢查檔案格式');
        } finally {
            hideLoadingOverlay();
        }
    });

    // 檔案上傳 - 參與者
    const participantFileUpload = document.getElementById('participantFileUpload');
    const participantFileInput = document.getElementById('participantFileInput');
    const uploadParticipantFileBtn = document.getElementById('uploadParticipantFile');
    const participantFileDisplay = document.getElementById('participantFileDisplay');

    participantFileUpload.addEventListener('click', () => {
        participantFileInput.click();
    });

    participantFileUpload.addEventListener('dragover', (e) => {
        e.preventDefault();
        participantFileUpload.classList.add('dragging');
    });

    participantFileUpload.addEventListener('dragleave', () => {
        participantFileUpload.classList.remove('dragging');
    });

    participantFileUpload.addEventListener('drop', (e) => {
        e.preventDefault();
        participantFileUpload.classList.remove('dragging');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            participantFileInput.files = files;
            displaySelectedFile(participantFileInput, participantFileDisplay, uploadParticipantFileBtn);
        }
    });

    participantFileInput.addEventListener('change', (e) => {
        displaySelectedFile(participantFileInput, participantFileDisplay, uploadParticipantFileBtn);
    });

    uploadParticipantFileBtn.addEventListener('click', async () => {
        const file = participantFileInput.files[0];

        if (!file) {
            showAlertModal('請選擇檔案');
            return;
        }

        if (!currentActivityId) {
            showAlertModal('無法確定活動ID');
            return;
        }

        try {
            showLoadingOverlay();
            const result = await handleParticipantFileUpload(file);

            if (result.errors.length > 0) {
                showAlertModal(`上傳完成，但有 ${result.errors.length} 筆資料有問題：\n${result.errors.join('\n')}`);
            }

            if (result.participants.length > 0) {
                const activity = activities.find(a => a.id === currentActivityId);
                if (activity) {
                    if (!activity.participants) {
                        activity.participants = [];
                    }

                    // 去重並新增參與者
                    const newParticipants = result.participants.filter(id => !activity.participants.includes(id));
                    activity.participants.push(...newParticipants);

                    const activityRef = window.firebase.ref(window.firebase.db, `activities/${currentActivityId}/participants`);
                    await window.firebase.set(activityRef, activity.participants);

                    renderParticipantsList();
                    renderActivitiesList();

                    showAlertModal(`成功新增 ${newParticipants.length} 位參與者`);
                }
            }

            clearSelectedFile(participantFileInput, participantFileDisplay, uploadParticipantFileBtn);

        } catch (error) {
            console.error('上傳參與者檔案錯誤:', error);
            showAlertModal('檔案上傳失敗，請檢查檔案格式');
        } finally {
            hideLoadingOverlay();
        }
    });

    // 編輯學生表單事件
    const cancelEditStudentBtn = document.getElementById('cancelEditStudent');
    if (cancelEditStudentBtn) {
        cancelEditStudentBtn.addEventListener('click', () => {
            hideModal('editStudentModal', 'editStudentBox');
        });
    }

    const editStudentForm = document.getElementById('editStudentForm');
    if (editStudentForm) {
        editStudentForm.addEventListener('submit', async (e) => {
            e.preventDefault();

            const studentId = document.getElementById('editStudentId').value;
            const formData = new FormData(e.target);

            const updatedData = {
                name: formData.get('editStudentName'),
                admissionYear: formData.get('editStudentYear'),
                updatedAt: new Date().toISOString()
            };

            await updateStudent(studentId, updatedData);
            hideModal('editStudentModal', 'editStudentBox');
        });
    }

    const unbindGoogleAccountBtn = document.getElementById('unbindGoogleAccount');
    if (unbindGoogleAccountBtn) {
        unbindGoogleAccountBtn.addEventListener('click', () => {
            const studentId = document.getElementById('editStudentId').value;
            if (studentId) {
                unbindGoogleAccount(studentId);
            }
        });
    }

    // 搜尋和篩選事件
    document.getElementById('activitySearchInput').addEventListener('input', filterActivities);
    document.getElementById('activityStatusFilter').addEventListener('change', filterActivities);
    document.getElementById('activityYearFilter').addEventListener('change', filterActivities);
    document.getElementById('studentSearchInput').addEventListener('input', filterStudents);
    document.getElementById('bindingStatusFilter').addEventListener('change', filterStudents);

    // 清除篩選事件
    document.getElementById('clearActivityFilters').addEventListener('click', clearActivityFilters);
    document.getElementById('clearStudentFilters').addEventListener('click', clearStudentFilters);

    // 全選事件
    document.getElementById('selectAllActivities').addEventListener('change', (e) => {
        const isChecked = e.target.checked;
        const displayedActivityIds = filteredActivities.map(a => a.id);

        if (isChecked) {
            displayedActivityIds.forEach(id => selectedActivities.add(id));
        } else {
            displayedActivityIds.forEach(id => selectedActivities.delete(id));
        }

        document.querySelectorAll('.activity-checkbox').forEach(cb => cb.checked = isChecked);
        updateActivityBulkActions();
    });

    document.getElementById('selectAllStudents').addEventListener('change', (e) => {
        const isChecked = e.target.checked;
        const displayedStudentIds = filteredStudents.map(s => s.id);

        if (isChecked) {
            displayedStudentIds.forEach(id => selectedStudents.add(id));
        } else {
            displayedStudentIds.forEach(id => selectedStudents.delete(id));
        }

        document.querySelectorAll('.student-checkbox').forEach(cb => cb.checked = isChecked);
        updateStudentBulkActions();
    });

    // 批量操作事件
    document.getElementById('bulkDeleteActivities').addEventListener('click', bulkDeleteActivities);
    document.getElementById('bulkToggleActivitiesVisibility').addEventListener('click', bulkToggleActivitiesVisibility);
    document.getElementById('cancelActivitySelection').addEventListener('click', clearActivitySelection);

    document.getElementById('bulkDeleteStudents').addEventListener('click', bulkDeleteStudents);
    document.getElementById('bulkUnbindStudents').addEventListener('click', bulkUnbindStudents);
    document.getElementById('cancelStudentSelection').addEventListener('click', clearStudentSelection);

    // 排序事件
    document.querySelectorAll('.activities-table .sortable-header').forEach(header => {
        header.addEventListener('click', () => {
            const field = header.dataset.sort;
            if (currentActivitySort.field === field) {
                currentActivitySort.direction = currentActivitySort.direction === 'asc' ? 'desc' : 'asc';
            } else {
                currentActivitySort.field = field;
                currentActivitySort.direction = 'asc';
            }
            applyActivitySort();
            renderActivitiesList();
            updateActivitySortHeaders();
        });
    });

    document.querySelectorAll('.students-table .sortable-header').forEach(header => {
        header.addEventListener('click', () => {
            const field = header.dataset.sort;
            if (currentStudentSort.field === field) {
                currentStudentSort.direction = currentStudentSort.direction === 'asc' ? 'desc' : 'asc';
            } else {
                currentStudentSort.field = field;
                currentStudentSort.direction = 'asc';
            }
            applyStudentSort();
            renderStudentsList();
            updateStudentSortHeaders();
        });
    });

    // 手動新增參與者
    document.getElementById('addParticipant').addEventListener('click', () => {
        const studentId = document.getElementById('participantStudentId').value.trim();
        if (studentId) {
            addParticipant(studentId);
            document.getElementById('participantStudentId').value = '';
        } else {
            showAlertModal('請輸入學號');
        }
    });

    // 參與者Modal關閉按鈕
    document.getElementById('cancelParticipants').addEventListener('click', () => {
        hideModal('participantsModal', 'participantsBox');
    });

    // 允許按Enter鍵新增參與者
    document.getElementById('participantStudentId').addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            const studentId = e.target.value.trim();
            if (studentId) {
                addParticipant(studentId);
                e.target.value = '';
            }
        }
    });
});

// 下載學生總表
function downloadStudentsSummary() {
    try {
        const yearFilter = document.getElementById('yearFilter').value;

        if (!yearFilter) {
            showAlertModal('請先選擇入學年度再下載總表');
            return;
        }

        // 根據選擇的入學年過濾學生
        const filteredStudents = students.filter(student => student.admissionYear === yearFilter);

        if (filteredStudents.length === 0) {
            showAlertModal(`${yearFilter} 學年度目前沒有學生資料可下載`);
            return;
        }

        // 準備表格資料
        const tableData = filteredStudents.map(student => {
            const participantCount = getStudentParticipantCount(student.id);
            const participatedActivities = activities
                .filter(activity => activity.participants && activity.participants.includes(student.id))
                .map(activity => activity.name)
                .join('；');

            return {
                '姓名': student.name,
                '學號': student.id,
                '入學年': student.admissionYear + '學年度',
                '參與場次': participantCount,
                '已參與活動': participatedActivities || '無',
                'Google帳號': student.googleAccount || '未綁定',
                '綁定狀態': student.googleAccount ? '已綁定' : '未綁定'
            };
        });

        // 建立工作簿
        const wb = XLSX.utils.book_new();

        // 建立工作表
        const ws = XLSX.utils.json_to_sheet(tableData);

        // 設定欄位寬度
        const columnWidths = [
            { wch: 10 }, // 姓名
            { wch: 12 }, // 學號
            { wch: 8 },  // 入學年
            { wch: 8 },  // 參與場次
            { wch: 50 }, // 已參與活動
            { wch: 25 }, // Google帳號
            { wch: 10 }  // 綁定狀態
        ];
        ws['!cols'] = columnWidths;

        // 將工作表加入工作簿
        XLSX.utils.book_append_sheet(wb, ws, `${yearFilter} 學年度學生總表`);

        // 產生檔案名稱（包含入學年和日期）
        const now = new Date();
        const dateStr = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
        const fileName = `${yearFilter} 學年度學生總表_${dateStr}.xlsx`;

        // 下載檔案
        XLSX.writeFile(wb, fileName);

        showAlertModal(`已下載 ${yearFilter} 學年度學生總表，共 ${filteredStudents.length} 筆資料`);
    } catch (error) {
        console.error('下載學生總表錯誤:', error);
        showAlertModal('下載失敗，請重試');
    }
}

// 處理學生檔案上傳
async function handleStudentFileUpload(file, admissionYear) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(sheet);

                const result = {
                    students: [],
                    errors: []
                };

                jsonData.forEach((row, index) => {
                    const rowNum = index + 2; // Excel 行號從 2 開始（第一行是標題）

                    // 尋找學號欄位（支援多種可能的欄位名稱）
                    const studentIdField = Object.keys(row).find(key =>
                        ['學號', '學生學號', 'ID', 'StudentID', 'Student ID'].includes(key.trim())
                    );

                    // 尋找姓名欄位（支援多種可能的欄位名稱）
                    const nameField = Object.keys(row).find(key =>
                        ['姓名', '學生姓名', 'Name', 'StudentName', 'Student Name'].includes(key.trim())
                    );

                    if (!studentIdField || !nameField) {
                        result.errors.push(`第 ${rowNum} 行：找不到學號或姓名欄位`);
                        return;
                    }

                    const studentId = String(row[studentIdField] || '').trim();
                    const name = String(row[nameField] || '').trim();

                    if (!studentId || !name) {
                        result.errors.push(`第 ${rowNum} 行：學號或姓名為空`);
                        return;
                    }

                    // 檢查學號是否已存在
                    if (students.find(s => s.id === studentId)) {
                        result.errors.push(`第 ${rowNum} 行：學號 ${studentId} 已存在`);
                        return;
                    }

                    // 檢查本次上傳是否有重複學號
                    if (result.students.find(s => s.id === studentId)) {
                        result.errors.push(`第 ${rowNum} 行：學號 ${studentId} 在檔案中重複`);
                        return;
                    }

                    result.students.push({
                        id: studentId,
                        name: name,
                        admissionYear: admissionYear,
                        createdAt: new Date().toISOString()
                    });
                });

                resolve(result);
            } catch (error) {
                reject(error);
            }
        };

        reader.onerror = () => reject(new Error('檔案讀取失敗'));
        reader.readAsArrayBuffer(file);
    });
}

// 處理參與者檔案上傳
async function handleParticipantFileUpload(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(sheet);

                const result = {
                    participants: [],
                    errors: []
                };

                jsonData.forEach((row, index) => {
                    const rowNum = index + 2; // Excel 行號從 2 開始

                    // 尋找學號欄位
                    const studentIdField = Object.keys(row).find(key =>
                        ['學號', '學生學號', 'ID', 'StudentID', 'Student ID'].includes(key.trim())
                    );

                    if (!studentIdField) {
                        result.errors.push(`第 ${rowNum} 行：找不到學號欄位`);
                        return;
                    }

                    const studentId = String(row[studentIdField] || '').trim();

                    if (!studentId) {
                        result.errors.push(`第 ${rowNum} 行：學號為空`);
                        return;
                    }

                    // 檢查學號是否存在於學生資料中
                    const student = students.find(s => s.id === studentId);
                    if (!student) {
                        result.errors.push(`第 ${rowNum} 行：找不到學號 ${studentId} 的學生資料`);
                        return;
                    }

                    // 檢查是否重複
                    if (!result.participants.includes(studentId)) {
                        result.participants.push(studentId);
                    }
                });

                resolve(result);
            } catch (error) {
                reject(error);
            }
        };

        reader.onerror = () => reject(new Error('檔案讀取失敗'));
        reader.readAsArrayBuffer(file);
    });
}

// 顯示活動 Modal (新增或編輯)
function showActivityModal(activityId = null) {
    editingActivityId = activityId;
    const form = document.getElementById('activityForm');
    const modalTitle = document.getElementById('activityTitle');
    const deleteButton = document.getElementById('deleteActivity');

    form.reset(); // 重設表單

    if (activityId) {
        // 編輯模式
        modalTitle.textContent = '編輯活動';
        deleteButton.classList.remove('hidden');
        const activity = activities.find(a => a.id === activityId);
        if (activity) {
            document.getElementById('activityName').value = activity.name;
            document.getElementById('activityDate').value = activity.date;
            document.getElementById('activityLocation').value = activity.location;
            document.getElementById('activityTeacher').value = activity.teacher;
            document.getElementById('activityNotes').value = activity.notes || '';
            document.getElementById('activityVisible').checked = activity.visible;
        }
    } else {
        // 新增模式
        modalTitle.textContent = '新增活動';
        deleteButton.classList.add('hidden');
        // 可以設定預設日期為今天
        document.getElementById('activityDate').valueAsDate = new Date();
    }
    showModal('activityModal', 'activityBox');
}

// 編輯活動 (由按鈕觸發)
function editActivity(activityId) {
    showActivityModal(activityId);
}

// 刪除活動
async function deleteActivity(activityId) {
    const confirmed = await showConfirmModal('確定要刪除此活動嗎？此操作無法撤銷。');
    if (!confirmed) return;

    try {
        showLoadingOverlay();
        const activityRef = window.firebase.ref(window.firebase.db, `activities/${activityId}`);
        await window.firebase.remove(activityRef);

        activities = activities.filter(a => a.id !== activityId);
        filterActivities(); // 更新列表
        showAlertModal('活動已刪除');
    } catch (error) {
        console.error('刪除活動錯誤:', error);
        showAlertModal('刪除活動失敗，請重試');
    } finally {
        hideLoadingOverlay();
    }
}

// 新增學年
async function addAdmissionYear(year) {
    if (!year || isNaN(parseInt(year))) {
        showAlertModal('請輸入有效的年份');
        return;
    }
    try {
        showLoadingOverlay();
        const yearRef = window.firebase.ref(window.firebase.db, `admissionYears/${year}`);
        await window.firebase.set(yearRef, { createdAt: new Date().toISOString() });

        admissionYears.push({ year: year, createdAt: new Date().toISOString() });
        admissionYears.sort((a, b) => parseInt(b.year) - parseInt(a.year)); // 重新排序
        updateYearSelects();
        showAlertModal(`入學年 ${year} 已新增`);

        hideModal('yearModal', 'yearBox');
        document.getElementById('yearForm').reset();
    } catch (error) {
        console.error('新增入學年錯誤:', error);
        showAlertModal('新增入學年失敗，請重試');
    } finally {
        hideLoadingOverlay();
    }
}

// 新增學生 (手動)
async function addStudent(studentData) {

    if (!studentData.id || !studentData.name || !studentData.admissionYear) {
        showAlertModal('學號、姓名和入學年為必填欄位');
        return;
    }

    // 檢查學號是否已存在
    if (students.find(s => s.id === studentData.id)) {
        showAlertModal(`學號 ${studentData.id} 已存在`);
        return;
    }

    try {
        showLoadingOverlay();
        const studentRef = window.firebase.ref(window.firebase.db, `students/${studentData.id}`);
        await window.firebase.set(studentRef, studentData);

        students.push(studentData);
        filterStudents(); // 更新列表
        showAlertModal(`學生 ${studentData.name} 已新增`);
        hideModal('studentModal', 'studentBox'); // 如果是從 Modal 新增，則關閉
    } catch (error) {
        console.error('新增學生錯誤:', error);
        showAlertModal('新增學生失敗，請重試');
    } finally {
        hideLoadingOverlay();
    }
}

// 編輯學生 (開啟 Modal)
function editStudent(studentId) {
    const student = students.find(s => s.id === studentId);
    if (!student) {
        showAlertModal('找不到學生資料');
        return;
    }

    document.getElementById('editStudentId').value = student.id; // 隱藏欄位，但用於表單提交
    document.getElementById('editStudentName').value = student.name;
    document.getElementById('editStudentYear').value = student.admissionYear;
    document.getElementById('editStudentGoogleAccount').value = student.googleAccount || '';

    const unbindBtn = document.getElementById('unbindGoogleAccount');
    if (student.googleAccount) {
        unbindBtn.classList.remove('hidden');
        unbindBtn.disabled = false;
    } else {
        unbindBtn.classList.add('hidden');
        unbindBtn.disabled = true;
    }

    showModal('editStudentModal', 'editStudentBox');
}

// 更新學生資料
async function updateStudent(studentId, updatedData) {
    if (!updatedData.name || !updatedData.admissionYear) {
        showAlertModal('姓名和入學年為必填欄位');
        return;
    }
    try {
        showLoadingOverlay();
        const studentRef = window.firebase.ref(window.firebase.db, `students/${studentId}`);
        await window.firebase.update(studentRef, updatedData);

        const index = students.findIndex(s => s.id === studentId);
        if (index !== -1) {
            students[index] = { ...students[index], ...updatedData };
        }
        filterStudents(); // 更新列表
        showAlertModal('學生資料已更新');
    } catch (error) {
        console.error('更新學生資料錯誤:', error);
        showAlertModal('更新學生資料失敗，請重試');
    } finally {
        hideLoadingOverlay();
    }
}

// 刪除學生
async function deleteStudent(studentId) {
    const confirmed = await showConfirmModal('確定要刪除此學生嗎？相關的活動參與紀錄也將被移除。此操作無法撤銷。');
    if (!confirmed) return;

    try {
        showLoadingOverlay();
        // 刪除學生資料
        const studentRef = window.firebase.ref(window.firebase.db, `students/${studentId}`);
        await window.firebase.remove(studentRef);

        // 從所有活動中移除該學生
        const activityUpdatePromises = activities.map(activity => {
            if (activity.participants && activity.participants.includes(studentId)) {
                const newParticipants = activity.participants.filter(id => id !== studentId);
                activity.participants = newParticipants; // 更新本地活動資料
                const activityParticipantsRef = window.firebase.ref(window.firebase.db, `activities/${activity.id}/participants`);
                return window.firebase.set(activityParticipantsRef, newParticipants);
            }
            return Promise.resolve();
        });
        await Promise.all(activityUpdatePromises);

        // 從 users 表中移除（如果已綁定）
        const student = students.find(s => s.id === studentId);
        if (student && student.googleAccount) {
            const userEmailKey = student.googleAccount.replace(/\./g, ',');
            const userRef = window.firebase.ref(window.firebase.db, `users/${userEmailKey}`);
            await window.firebase.remove(userRef);
        }

        students = students.filter(s => s.id !== studentId);
        filterStudents(); // 更新列表
        showAlertModal('學生已刪除');
    } catch (error) {
        console.error('刪除學生錯誤:', error);
        showAlertModal('刪除學生失敗，請重試');
    } finally {
        hideLoadingOverlay();
    }
}

// 解除 Google 帳號綁定
async function unbindGoogleAccount(studentId) {
    const student = students.find(s => s.id === studentId);
    if (!student || !student.googleAccount) {
        showAlertModal('此學生未綁定Google帳號');
        return;
    }

    const confirmed = await showConfirmModal(`確定要解除學生 ${student.name} (${student.id}) 的Google帳號 (${student.googleAccount}) 綁定嗎？`);
    if (!confirmed) return;

    try {
        showLoadingOverlay();
        // 從學生資料中移除 googleAccount
        const studentGoogleRef = window.firebase.ref(window.firebase.db, `students/${studentId}/googleAccount`);
        await window.firebase.remove(studentGoogleRef);

        // 從 users 表中移除該使用者
        const userEmailKey = student.googleAccount.replace(/\./g, ',');
        const userRef = window.firebase.ref(window.firebase.db, `users/${userEmailKey}`);
        await window.firebase.remove(userRef);

        // 更新本地資料
        delete student.googleAccount;
        filterStudents(); // 更新列表
        hideModal('editStudentModal', 'editStudentBox');
        showAlertModal('Google帳號已解除綁定');
    } catch (error) {
        console.error('解除綁定錯誤:', error);
        showAlertModal('解除綁定失敗，請重試');
    } finally {
        hideLoadingOverlay();
    }
}

// 獲取學生參與活動次數
function getStudentParticipantCount(studentId) {
    return activities.filter(activity => activity.participants && activity.participants.includes(studentId)).length;
}

// 渲染學生介面的活動列表
function renderStudentActivitiesList() {
    const listContainer = document.getElementById('studentActivitiesList');
    const totalActivitiesElement = document.getElementById('totalActivities');
    listContainer.innerHTML = ''; // 清空列表

    if (!studentData || !studentData.id) {
        listContainer.innerHTML = '<p class="text-gray-400">無法載入您的活動紀錄。</p>';
        totalActivitiesElement.textContent = '0';
        return;
    }

    const studentActivities = activities.filter(activity =>
        activity.visible && activity.participants && activity.participants.includes(studentData.id)
    ).sort((a, b) => new Date(b.date) - new Date(a.date)); // 按日期降序排序

    totalActivitiesElement.textContent = studentActivities.length;

    if (studentActivities.length === 0) {
        listContainer.innerHTML = '<p class="text-gray-400">目前還沒有任何活動紀錄。如果有任何問題，請聯絡所辦。</p>';
        return;
    }

    studentActivities.forEach(activity => {
        const card = document.createElement('div');
        card.className = 'activity-card p-4 rounded-lg shadow bg-dark-200 border border-dark-300';
        card.innerHTML = `
            <h3 class="text-lg font-semibold text-gary-100 mb-2">${activity.name}</h3>
            <p class="text-sm text-gray-400 mb-1">日期：${formatDate(activity.date)}</p>
            <p class="text-sm text-gray-400 mb-1">地點：${activity.location}</p>
            <p class="text-sm text-gray-400">承辦教師：${activity.teacher}</p>
            ${activity.notes ? `<p class="text-xs text-gray-500 mt-2">備註：${activity.notes}</p>` : ''}
        `;
        listContainer.appendChild(card);
    });
}

// 管理參與學生 (開啟 Modal)
function manageParticipants(activityId) {
    currentActivityId = activityId;
    const activity = activities.find(a => a.id === activityId);
    if (!activity) {
        showAlertModal('找不到活動資料');
        return;
    }
    document.querySelector('#participantsModal h3').textContent = `管理參與學生 - ${activity.name}`;
    renderParticipantsList();
    showModal('participantsModal', 'participantsBox');
}

// 渲染參與學生列表
function renderParticipantsList() {
    const listElement = document.getElementById('participantsList');
    listElement.innerHTML = '';

    if (!currentActivityId) return;
    const activity = activities.find(a => a.id === currentActivityId);
    if (!activity || !activity.participants || activity.participants.length === 0) {
        listElement.innerHTML = '<p class="text-gray-400 text-sm">此活動尚無參與學生。</p>';
        return;
    }

    activity.participants.forEach(studentId => {
        const student = students.find(s => s.id === studentId);
        const participantName = student ? `${student.name} (${studentId})` : `未知學生 (${studentId})`;

        const item = document.createElement('div');
        item.className = 'flex justify-between items-center p-2 bg-dark-200 rounded mb-1';
        item.innerHTML = `
            <span class="text-sm text-gray-300">${participantName}</span>
            <button class="remove-participant-btn text-red-400 hover:text-red-300 text-xs" data-student-id="${studentId}">移除</button>
        `;
        listElement.appendChild(item);
    });

    // 綁定移除按鈕事件
    listElement.querySelectorAll('.remove-participant-btn').forEach(button => {
        button.addEventListener('click', (e) => {
            const studentIdToRemove = e.target.dataset.studentId;
            removeParticipant(studentIdToRemove);
        });
    });
}

// 新增參與學生
async function addParticipant(studentId) {
    if (!currentActivityId || !studentId) return;

    const student = students.find(s => s.id === studentId);
    if (!student) {
        showAlertModal(`找不到學號為 ${studentId} 的學生資料。`);
        return;
    }

    const activity = activities.find(a => a.id === currentActivityId);
    if (!activity) return;

    if (!activity.participants) {
        activity.participants = [];
    }

    if (activity.participants.includes(studentId)) {
        showAlertModal(`學生 ${student.name} (${studentId}) 已在此活動的參與名單中。`);
        return;
    }

    try {
        showLoadingOverlay();
        activity.participants.push(studentId);
        const activityRef = window.firebase.ref(window.firebase.db, `activities/${currentActivityId}/participants`);
        await window.firebase.set(activityRef, activity.participants);

        renderParticipantsList();
        renderActivitiesList(); // 更新活動列表中的參與人數
        showAlertModal(`已新增學生 ${student.name} (${studentId}) 至活動。`);
    } catch (error) {
        console.error('新增參與者錯誤:', error);
        showAlertModal('新增參與者失敗，請重試');
        // 如果失敗，從本地 participants 陣列中移除
        activity.participants = activity.participants.filter(id => id !== studentId);
    } finally {
        hideLoadingOverlay();
    }
}

// 移除參與學生
async function removeParticipant(studentId) {
    if (!currentActivityId || !studentId) return;

    const activity = activities.find(a => a.id === currentActivityId);
    if (!activity || !activity.participants) return;

    const confirmed = await showConfirmModal(`確定要從此活動移除學生 (${studentId}) 嗎？`);
    if (!confirmed) return;

    try {
        showLoadingOverlay();
        activity.participants = activity.participants.filter(id => id !== studentId);
        const activityRef = window.firebase.ref(window.firebase.db, `activities/${currentActivityId}/participants`);
        await window.firebase.set(activityRef, activity.participants);

        renderParticipantsList();
        renderActivitiesList();
        showAlertModal(`已從活動移除學生 (${studentId})。`);
    } catch (error) {
        console.error('移除參與者錯誤:', error);
        showAlertModal('移除參與者失敗，請重試');
    } finally {
        hideLoadingOverlay();
    }
}

// 顯示選擇的檔案
function displaySelectedFile(fileInput, displayContainer, uploadButton) {
    const file = fileInput.files[0];

    if (file) {
        const fileSize = formatFileSize(file.size);
        const fileName = file.name;

        displayContainer.innerHTML = `
            <div class="selected-file">
                <div class="flex items-center justify-between p-3 bg-dark-200 border border-dark-300 rounded-md">
                    <div class="flex items-center gap-3">
                        <svg class="w-6 h-6 text-green-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" 
                                  d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z">
                            </path>
                        </svg>
                        <div>
                            <div class="text-sm font-medium text-gray-200">${fileName}</div>
                            <div class="text-xs text-gray-400">${fileSize}</div>
                        </div>
                    </div>
                    <button type="button" class="remove-file-btn text-red-400 hover:text-red-300 p-1 rounded hover:bg-red-400/10 transition-colors" title="移除檔案">
                        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path>
                        </svg>
                    </button>
                </div>
            </div>
        `;

        // 綁定移除檔案事件
        const removeBtn = displayContainer.querySelector('.remove-file-btn');
        removeBtn.addEventListener('click', () => {
            clearSelectedFile(fileInput, displayContainer, uploadButton);
        });

        uploadButton.disabled = false;
        displayContainer.classList.remove('hidden');
    } else {
        clearSelectedFile(fileInput, displayContainer, uploadButton);
    }
}

// 清除選擇的檔案
function clearSelectedFile(fileInput, displayContainer, uploadButton) {
    fileInput.value = '';
    displayContainer.innerHTML = '';
    displayContainer.classList.add('hidden');
    uploadButton.disabled = true;
}

// 格式化檔案大小
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}