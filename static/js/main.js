// App State
let leads = [];
let activeFilter = 'all';
let searchQuery = '';

// DOM Elements
const leadsTbody = document.getElementById('leads-tbody');
const searchInput = document.getElementById('search-input');
const tabButtons = document.querySelectorAll('.tab-btn');
const inputSheetUrl = document.getElementById('input-sheet-url');
const inputApiKey = document.getElementById('input-api-key');
const btnToggleKey = document.getElementById('btn-toggle-key');
const btnFetchSheet = document.getElementById('btn-fetch-sheet');
const btnLoadMock = document.getElementById('btn-load-mock');
const btnRunScoring = document.getElementById('btn-run-scoring');
const btnExportExcel = document.getElementById('btn-export-excel');
const inputFileUpload = document.getElementById('input-file-upload');

// Modal Elements
const leadModal = document.getElementById('lead-modal');
const closeModalBtns = document.querySelectorAll('.btn-close-modal');
const editLeadForm = document.getElementById('edit-lead-form');
const btnSaveLead = document.getElementById('btn-save-lead');
const btnDecisions = document.querySelectorAll('.btn-decision');

// Statistics Elements
const statVip = document.getElementById('stat-vip-count');
const statMedium = document.getElementById('stat-medium-count');
const statTrash = document.getElementById('stat-trash-count');
const statApproved = document.getElementById('stat-approved-count');

// Tab Counter Elements
const countAll = document.getElementById('count-all');
const countVip = document.getElementById('count-vip');
const countMedium = document.getElementById('count-medium');
const countTrash = document.getElementById('count-trash');
const countPending = document.getElementById('count-pending');

// Notification Elements
const toast = document.getElementById('toast');
const toastMessage = document.getElementById('toast-message');

// API Key Visibility Toggle
btnToggleKey.addEventListener('click', () => {
    const type = inputApiKey.getAttribute('type') === 'password' ? 'text' : 'password';
    inputApiKey.setAttribute('type', type);
    const icon = btnToggleKey.querySelector('i');
    if (type === 'text') {
        icon.setAttribute('data-lucide', 'eye-off');
    } else {
        icon.setAttribute('data-lucide', 'eye');
    }
    lucide.createIcons();
});

// Toast Helper
function showToast(message, type = 'info') {
    toastMessage.textContent = message;
    const icon = toast.querySelector('.toast-icon i');
    
    if (type === 'success') {
        toast.style.borderColor = 'var(--emerald)';
        icon.setAttribute('data-lucide', 'check-circle-2');
        icon.style.color = 'var(--emerald)';
    } else if (type === 'error') {
        toast.style.borderColor = 'var(--rose)';
        icon.setAttribute('data-lucide', 'alert-circle');
        icon.style.color = 'var(--rose)';
    } else {
        toast.style.borderColor = 'var(--primary)';
        icon.setAttribute('data-lucide', 'info');
        icon.style.color = 'var(--primary)';
    }
    
    lucide.createIcons();
    toast.classList.add('show');
    
    setTimeout(() => {
        toast.classList.remove('show');
    }, 4000);
}

// Initial Data Load
async function fetchLeads() {
    try {
        const response = await fetch('/api/leads');
        leads = await response.json();
        renderDashboard();
    } catch (error) {
        console.error('Error fetching leads:', error);
        showToast('Không thể kết nối đến backend API', 'error');
    }
}

// Render Dashboard (Stats, Tabs, Table)
function renderDashboard() {
    // 1. Calculate Statistics
    let vipCount = 0;
    let medCount = 0;
    let trashCount = 0;
    let appCount = 0;
    let pendingCount = 0;
    
    leads.forEach(lead => {
        if (lead.classification === 'VIP') vipCount++;
        else if (lead.classification === 'Trung bình') medCount++;
        else if (lead.classification === 'Rác') trashCount++;
        
        if (lead.status === 'Đã duyệt') appCount++;
        else if (lead.status === 'Chưa duyệt') pendingCount++;
    });
    
    // Update Stats cards
    statVip.textContent = vipCount;
    statMedium.textContent = medCount;
    statTrash.textContent = trashCount;
    statApproved.textContent = appCount;
    
    // Update Tab count labels
    countAll.textContent = leads.length;
    countVip.textContent = vipCount;
    countMedium.textContent = medCount;
    countTrash.textContent = trashCount;
    countPending.textContent = pendingCount;
    
    // 2. Filter & Search Leads
    let filteredLeads = leads;
    
    // Filter Tab
    if (activeFilter === 'VIP') {
        filteredLeads = leads.filter(l => l.classification === 'VIP');
    } else if (activeFilter === 'Trung bình') {
        filteredLeads = leads.filter(l => l.classification === 'Trung bình');
    } else if (activeFilter === 'Rác') {
        filteredLeads = leads.filter(l => l.classification === 'Rác');
    } else if (activeFilter === 'Chưa duyệt') {
        filteredLeads = leads.filter(l => l.status === 'Chưa duyệt');
    }
    
    // Search query
    if (searchQuery.trim() !== '') {
        const query = searchQuery.toLowerCase().trim();
        filteredLeads = filteredLeads.filter(l => {
            const name = (l.name || '').toLowerCase();
            const phone = (l.phone || '').toLowerCase();
            const email = (l.email || '').toLowerCase();
            const req = (l.requirement || '').toLowerCase();
            const reason = (l.reason || '').toLowerCase();
            return name.includes(query) || phone.includes(query) || email.includes(query) || req.includes(query) || reason.includes(query);
        });
    }
    
    // 3. Render Table rows
    if (filteredLeads.length === 0) {
        leadsTbody.innerHTML = `
            <tr>
                <td colspan="8" class="text-center py-8">
                    <div class="loading-state">
                        <i data-lucide="inbox"></i>
                        <p class="mt-2 text-muted">Không tìm thấy khách hàng nào khớp bộ lọc.</p>
                    </div>
                </td>
            </tr>
        `;
        lucide.createIcons();
        return;
    }
    
    leadsTbody.innerHTML = filteredLeads.map(lead => {
        // AI Score badge
        let scoreClass = 'neutral';
        let scoreText = lead.score !== null ? lead.score : '--';
        if (lead.score > 0) scoreClass = 'positive';
        else if (lead.score < 0) scoreClass = 'negative';
        
        // AI Classification tag
        let classClass = 'medium';
        let classText = lead.classification || 'Chưa có';
        if (lead.classification === 'VIP') classClass = 'vip';
        else if (lead.classification === 'Rác') classClass = 'trash';
        
        // Human Status badge
        let statusClass = 'pending';
        if (lead.status === 'Đã duyệt') statusClass = 'approved';
        else if (lead.status === 'Từ chối') statusClass = 'rejected';
        
        const leadCode = `L${String(lead.id).padStart(3, '0')}`;
        
        return `
            <tr onclick="openEditModal(${lead.id})">
                <td><span class="lead-code">${leadCode}</span></td>
                <td>
                    <div class="lead-name-cell">
                        <span class="lead-name">${escapeHTML(lead.name)}</span>
                        ${lead.email ? `<span class="lead-email">${escapeHTML(lead.email)}</span>` : ''}
                    </div>
                </td>
                <td>${escapeHTML(lead.phone || 'Không có')}</td>
                <td>
                    <div class="lead-req-excerpt" title="${escapeHTML(lead.requirement)}">
                        ${escapeHTML(lead.requirement || 'Chưa cung cấp nhu cầu')}
                    </div>
                </td>
                <td class="text-center">
                    <span class="score-badge ${scoreClass}">${scoreText}</span>
                </td>
                <td>
                    <span class="class-tag ${classClass}">${classText}</span>
                </td>
                <td>
                    <span class="status-badge ${statusClass}">${lead.status}</span>
                </td>
                <td class="text-center" onclick="event.stopPropagation()">
                    <button class="btn-action" onclick="openEditModal(${lead.id})">
                        <i data-lucide="edit-3"></i>
                    </button>
                </td>
            </tr>
        `;
    }).join('');
    
    lucide.createIcons();
}

// Escape HTML utility
function escapeHTML(str) {
    if (!str) return '';
    return str.replace(/[&<>'"]/g, 
        tag => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[tag] || tag)
    );
}

// Filter Tab Click Handlers
tabButtons.forEach(button => {
    button.addEventListener('click', () => {
        tabButtons.forEach(btn => btn.classList.remove('active'));
        button.classList.add('active');
        activeFilter = button.getAttribute('data-filter');
        renderDashboard();
    });
});

// Search Input Handler
searchInput.addEventListener('input', (e) => {
    searchQuery = e.target.value;
    renderDashboard();
});

// Google Sheet Fetch Click Handler
btnFetchSheet.addEventListener('click', async () => {
    const url = inputSheetUrl.value.trim();
    if (!url) {
        showToast('Vui lòng nhập đường dẫn Google Sheets', 'error');
        return;
    }
    
    btnFetchSheet.disabled = true;
    btnFetchSheet.innerHTML = '<i class="animate-spin" data-lucide="loader-2"></i><span>Đang kéo...</span>';
    lucide.createIcons();
    
    try {
        const response = await fetch('/api/fetch-sheet', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ url: url })
        });
        
        const result = await response.json();
        if (response.ok && result.status === 'success') {
            leads = result.data;
            renderDashboard();
            showToast(result.msg, 'success');
        } else {
            showToast(result.msg || 'Đồng bộ thất bại', 'error');
        }
    } catch (error) {
        console.error('Error fetching sheet:', error);
        showToast('Lỗi kết nối khi đồng bộ dữ liệu', 'error');
    } finally {
        btnFetchSheet.disabled = false;
        btnFetchSheet.innerHTML = '<i data-lucide="refresh-cw"></i><span>Đồng Bộ</span>';
        lucide.createIcons();
    }
});

// Load Mock Data Click Handler
btnLoadMock.addEventListener('click', async () => {
    btnLoadMock.disabled = true;
    try {
        const response = await fetch('/api/load-mock', { method: 'POST' });
        const result = await response.json();
        if (response.ok && result.status === 'success') {
            leads = result.data;
            renderDashboard();
            showToast(result.msg, 'success');
        } else {
            showToast(result.msg || 'Nạp dữ liệu mẫu thất bại', 'error');
        }
    } catch (error) {
        console.error('Error loading mock:', error);
        showToast('Lỗi khi nạp dữ liệu mẫu', 'error');
    } finally {
        btnLoadMock.disabled = false;
    }
});

// File Upload Handler
inputFileUpload.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    const formData = new FormData();
    formData.append('file', file);
    
    showToast('Đang phân tích tệp...', 'info');
    
    try {
        const response = await fetch('/api/upload-file', {
            method: 'POST',
            body: formData
        });
        const result = await response.json();
        if (response.ok && result.status === 'success') {
            leads = result.data;
            renderDashboard();
            showToast(result.msg, 'success');
        } else {
            showToast(result.msg || 'Không nạp được tệp', 'error');
        }
    } catch (error) {
        console.error('Error uploading file:', error);
        showToast('Lỗi kết nối khi tải tệp lên', 'error');
    } finally {
        inputFileUpload.value = ''; // reset file input
    }
});

// Run AI Scoring Handler
btnRunScoring.addEventListener('click', async () => {
    const apiKey = inputApiKey.value.trim();
    
    btnRunScoring.disabled = true;
    btnRunScoring.innerHTML = '<i class="animate-spin" data-lucide="loader-2"></i><span>Đang chấm điểm...</span>';
    lucide.createIcons();
    
    try {
        const response = await fetch('/api/score', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ api_key: apiKey })
        });
        
        const result = await response.json();
        if (response.ok && result.status === 'success') {
            leads = result.data;
            renderDashboard();
            if (result.warnings && result.warnings.length > 0) {
                showToast(result.msg, 'error');
            } else {
                showToast(result.msg, 'success');
            }
        } else {
            showToast(result.msg || 'Chấm điểm thất bại', 'error');
        }
    } catch (error) {
        console.error('Error running scoring:', error);
        showToast('Lỗi kết nối khi chấm điểm AI', 'error');
    } finally {
        btnRunScoring.disabled = false;
        btnRunScoring.innerHTML = '<i data-lucide="sparkles"></i><span>Kích Hoạt Chấm Điểm AI</span>';
        lucide.createIcons();
    }
});

// Export Excel Click Handler
btnExportExcel.addEventListener('click', () => {
    if (leads.length === 0) {
        showToast('Không có dữ liệu khách hàng để xuất Excel', 'error');
        return;
    }
    showToast('Đang khởi tạo file Excel...', 'info');
    window.location.href = '/api/export';
});

// --- MODAL CONTROLS ---

let currentModalLeadStatus = 'Chưa duyệt';

function openEditModal(id) {
    const lead = leads.find(l => l.id === id);
    if (!lead) return;
    
    // Set field values
    document.getElementById('edit-lead-id').value = lead.id;
    document.getElementById('edit-lead-name').value = lead.name || '';
    document.getElementById('edit-lead-phone').value = lead.phone || '';
    document.getElementById('edit-lead-email').value = lead.email || '';
    document.getElementById('edit-lead-req').value = lead.requirement || '';
    
    document.getElementById('edit-lead-score').value = lead.score !== null ? lead.score : '';
    document.getElementById('edit-lead-class').value = lead.classification || 'Trung bình';
    document.getElementById('edit-lead-reason').value = lead.reason || '';
    document.getElementById('edit-lead-reviewer-notes').value = lead.reviewer_notes || '';
    
    // Update header labels
    document.getElementById('modal-lead-code').textContent = `L${String(lead.id).padStart(3, '0')}`;
    document.getElementById('modal-lead-name-title').textContent = lead.name || 'Chi Tiết Khách Hàng';
    
    // Set active status decision button
    currentModalLeadStatus = lead.status || 'Chưa duyệt';
    updateDecisionButtonsUI(currentModalLeadStatus);
    
    // Open modal
    leadModal.classList.add('open');
}

function updateDecisionButtonsUI(status) {
    btnDecisions.forEach(btn => {
        const btnStatus = btn.getAttribute('data-status');
        if (btnStatus === status) {
            btn.classList.add('active');
        } else {
            btn.classList.remove('active');
        }
    });
}

// Modal Decision Buttons click handler
btnDecisions.forEach(btn => {
    btn.addEventListener('click', () => {
        const status = btn.getAttribute('data-status');
        currentModalLeadStatus = status;
        updateDecisionButtonsUI(status);
    });
});

// Close Modal functions
function closeModal() {
    leadModal.classList.remove('open');
}

closeModalBtns.forEach(btn => {
    btn.addEventListener('click', closeModal);
});

// Save Lead Handler
btnSaveLead.addEventListener('click', async () => {
    const id = parseInt(document.getElementById('edit-lead-id').value);
    const name = document.getElementById('edit-lead-name').value.trim();
    const phone = document.getElementById('edit-lead-phone').value.trim();
    const email = document.getElementById('edit-lead-email').value.trim();
    const requirement = document.getElementById('edit-lead-req').value.trim();
    
    const scoreVal = document.getElementById('edit-lead-score').value;
    const score = scoreVal !== '' ? parseInt(scoreVal) : null;
    const classification = document.getElementById('edit-lead-class').value;
    const reason = document.getElementById('edit-lead-reason').value.trim();
    const reviewer_notes = document.getElementById('edit-lead-reviewer-notes').value.trim();
    
    if (!name || !requirement) {
        showToast('Họ tên và chi tiết nhu cầu không được để trống', 'error');
        return;
    }
    
    const updatePayload = {
        id, name, phone, email, requirement,
        score, classification, reason,
        status: currentModalLeadStatus,
        reviewer_notes
    };
    
    btnSaveLead.disabled = true;
    btnSaveLead.innerHTML = '<i class="animate-spin" data-lucide="loader-2"></i><span>Đang lưu...</span>';
    lucide.createIcons();
    
    try {
        const response = await fetch('/api/leads/update', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(updatePayload)
        });
        
        const result = await response.json();
        if (response.ok && result.status === 'success') {
            leads = result.data;
            renderDashboard();
            closeModal();
            showToast(result.msg, 'success');
        } else {
            showToast(result.msg || 'Không thể lưu thay đổi', 'error');
        }
    } catch (error) {
        console.error('Error saving lead:', error);
        showToast('Lỗi kết nối khi cập nhật thông tin', 'error');
    } finally {
        btnSaveLead.disabled = false;
        btnSaveLead.innerHTML = '<i data-lucide="save"></i><span>Lưu Thay Đổi</span>';
        lucide.createIcons();
    }
});

// Load Initial Data on start
window.addEventListener('DOMContentLoaded', () => {
    fetchLeads();
});
