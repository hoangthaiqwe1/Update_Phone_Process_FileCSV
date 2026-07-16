// ============================================================
// APP.JS - Logic giao diện chính
// Xử lý 2 tính năng:
//   1. Tab Process: chọn file, validate, gửi lên server, hiển thị kết quả
//   2. Tab Update: chọn ticket, upload Result mới, cập nhật
// ============================================================

document.addEventListener('DOMContentLoaded', function () {

    // ===== BIẾN DOM - TAB PROCESS =====
    const fileCountSelect = document.getElementById('fileCount');
    const filePairsContainer = document.getElementById('filePairs');
    const processBtn = document.getElementById('processBtn');
    const progressContainer = document.getElementById('progressContainer');
    const progressBar = document.getElementById('progressBar');
    const resultsContainer = document.getElementById('resultsContainer');
    const resultsBody = document.getElementById('resultsBody');
    const summaryContainer = document.getElementById('summaryContainer');
    const downloadAllContainer = document.getElementById('downloadAllContainer');
    const warningsContainer = document.getElementById('warningsContainer');

    // ===== BIẾN DOM - TAB UPDATE =====
    const updateCountSelect = document.getElementById('updateCount');
    const updatePairs = document.getElementById('updatePairs');
    const updateBtn = document.getElementById('updateBtn');
    const updateResults = document.getElementById('updateResults');
    const updateResultsBody = document.getElementById('updateResultsBody');
    const updateSummary = document.getElementById('updateSummary');
    const previewTicketSelect = document.getElementById('previewTicketSelect');
    const previewBody = document.getElementById('previewBody');
    const previewCount = document.getElementById('previewCount');

    // Danh sách ticket đã xử lý (load từ server)
    let allTickets = [];

    /**
     * Tạo ô input có autocomplete gợi ý ticket
     * Khi gõ → lọc danh sách ticket khớp → hiện dropdown gợi ý
     */
    function createTicketSearchInput(name, placeholder) {
        return `
            <div class="position-relative ticket-search-wrapper">
                <input type="text" class="form-control form-control-sm ticket-search-input"
                       name="${name}" placeholder="${placeholder || 'Gõ tên ticket...'}"
                       autocomplete="off">
                <div class="ticket-suggestions" style="display:none;position:absolute;top:100%;left:0;right:0;
                    z-index:1000;max-height:200px;overflow-y:auto;background:#fff;
                    border:1px solid #dee2e6;border-radius:6px;box-shadow:0 4px 12px rgba(0,0,0,0.1);">
                </div>
            </div>
        `;
    }

    /**
     * Gắn sự kiện autocomplete cho tất cả ô ticket-search-input
     */
    function initTicketAutocomplete() {
        document.querySelectorAll('.ticket-search-input').forEach(input => {
            const wrapper = input.closest('.ticket-search-wrapper');
            const suggestions = wrapper.querySelector('.ticket-suggestions');

            // Gõ → lọc và hiện gợi ý
            input.addEventListener('input', function() {
                const query = this.value.trim().toLowerCase();
                if (!query) { suggestions.style.display = 'none'; return; }

                const filtered = allTickets.filter(t => t.toLowerCase().includes(query)).slice(0, 20);
                if (filtered.length === 0) {
                    suggestions.style.display = 'none';
                    return;
                }

                suggestions.innerHTML = filtered.map(t =>
                    `<div class="ticket-suggest-item" style="padding:8px 12px;cursor:pointer;font-size:0.85rem;
                        border-bottom:1px solid #f0f0f0;">${t}</div>`
                ).join('');
                suggestions.style.display = 'block';

                // Click vào gợi ý → chọn
                suggestions.querySelectorAll('.ticket-suggest-item').forEach(item => {
                    item.addEventListener('click', function() {
                        input.value = this.textContent;
                        suggestions.style.display = 'none';
                        input.dispatchEvent(new Event('change'));
                    });
                    // Hover effect
                    item.addEventListener('mouseenter', function() {
                        this.style.background = '#e9ecef';
                    });
                    item.addEventListener('mouseleave', function() {
                        this.style.background = '#fff';
                    });
                });
            });

            // Ẩn gợi ý khi click ra ngoài
            document.addEventListener('click', function(e) {
                if (!wrapper.contains(e.target)) {
                    suggestions.style.display = 'none';
                }
            });
        });
    }

    // ============================================================
    // TAB PROCESS - Tạo form upload khi chọn số lượng file
    // ============================================================
    if (fileCountSelect) {
        fileCountSelect.addEventListener('change', function () {
            generateFilePairs(parseInt(this.value));
        });
    }

    /**
     * Tạo các ô upload cho mỗi cặp file
     * Mỗi cặp gồm: Tên export + File D2C + File Result
     */
    function generateFilePairs(count) {
        filePairsContainer.innerHTML = '';
        if (count === 0) return;

        for (let i = 1; i <= count; i++) {
            const pair = document.createElement('div');
            pair.className = 'file-pair';
            pair.innerHTML = `
                <div class="file-pair-header">
                    <i class="bi bi-folder2-open"></i> Cặp file ${i}
                </div>
                <div class="row">
                    <div class="col-md-4 mb-2">
                        <label class="form-label small fw-semibold">Tên export ${i}</label>
                        <input type="text" class="form-control form-control-sm"
                               name="export_name_${i}" placeholder="VD: S-74155371">
                    </div>
                    <div class="col-md-4 mb-2">
                        <label class="form-label small fw-semibold">File D2C CS ${i}</label>
                        <input type="file" class="form-control form-control-sm"
                               name="d2c_${i}" accept=".xlsx,.xls,.csv" required>
                    </div>
                    <div class="col-md-4 mb-2">
                        <label class="form-label small fw-semibold">File Result FEOL ${i}</label>
                        <input type="file" class="form-control form-control-sm"
                               name="result_${i}" accept=".xlsx,.xls,.csv" required>
                    </div>
                </div>
            `;
            filePairsContainer.appendChild(pair);
        }
    }

    // ============================================================
    // TAB PROCESS - Nút PROCESS: validate + gửi lên server
    // ============================================================
    if (processBtn) {
        processBtn.addEventListener('click', function () {
            const fileCount = parseInt(fileCountSelect.value);
            if (fileCount === 0) {
                showAlert('Vui lòng chọn số lượng file!', 'warning');
                return;
            }

            // --- Validate: kiểm tra đủ thông tin ---
            const exportNames = [];
            for (let i = 1; i <= fileCount; i++) {
                const exportName = document.querySelector(`input[name="export_name_${i}"]`);
                const d2c = document.querySelector(`input[name="d2c_${i}"]`);
                const result = document.querySelector(`input[name="result_${i}"]`);

                // Phải có tên export
                if (!exportName || !exportName.value.trim()) {
                    showAlert(`Chưa đặt tên file export số ${i}`, 'warning');
                    if (exportName) exportName.focus();
                    return;
                }
                // Tên export không được trùng nhau
                const nameVal = exportName.value.trim();
                if (exportNames.includes(nameVal)) {
                    showAlert(`Tên export "${nameVal}" bị trùng! Đặt tên khác cho file ${i}`, 'danger');
                    exportName.focus();
                    return;
                }
                exportNames.push(nameVal);

                // Phải có file D2C
                if (!d2c || !d2c.files.length) {
                    showAlert(`Thiếu file D2C số ${i}`, 'danger');
                    return;
                }
                // Phải có file Result
                if (!result || !result.files.length) {
                    showAlert(`Thiếu file Result số ${i}`, 'danger');
                    return;
                }
            }

            // --- Đóng gói dữ liệu gửi lên server ---
            const formData = new FormData();
            formData.append('file_count', fileCount);
            formData.append('password', document.getElementById('password').value);

            for (let i = 1; i <= fileCount; i++) {
                const d2c = document.querySelector(`input[name="d2c_${i}"]`);
                const result = document.querySelector(`input[name="result_${i}"]`);
                const exportName = document.querySelector(`input[name="export_name_${i}"]`);
                formData.append(`d2c_${i}`, d2c.files[0]);
                formData.append(`result_${i}`, result.files[0]);
                formData.append(`export_name_${i}`, exportName.value.trim());
            }

            // --- Hiển thị progress bar ---
            progressContainer.classList.add('show');
            progressBar.style.width = '0%';
            progressBar.textContent = '0%';
            resultsContainer.style.display = 'none';
            summaryContainer.style.display = 'none';
            downloadAllContainer.style.display = 'none';
            warningsContainer.innerHTML = '';
            processBtn.disabled = true;
            processBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-2"></span>Đang xử lý...';

            // Giả lập progress
            let progress = 0;
            const progressInterval = setInterval(() => {
                if (progress < 90) {
                    progress += Math.random() * 15;
                    progress = Math.min(progress, 90);
                    progressBar.style.width = progress + '%';
                    progressBar.textContent = Math.round(progress) + '%';
                }
            }, 300);

            // --- Gửi request lên server ---
            fetch('/api/process', { method: 'POST', body: formData })
                .then(r => r.json())
                .then(data => {
                    clearInterval(progressInterval);
                    progressBar.style.width = '100%';
                    progressBar.textContent = '100%';
                    if (data.success) displayResults(data);
                    else showAlert(data.error || 'Có lỗi xảy ra', 'danger');
                })
                .catch(err => {
                    clearInterval(progressInterval);
                    showAlert('Lỗi kết nối: ' + err.message, 'danger');
                })
                .finally(() => {
                    processBtn.disabled = false;
                    processBtn.innerHTML = '<i class="bi bi-play-fill me-2"></i>PROCESS';
                    setTimeout(() => progressContainer.classList.remove('show'), 1000);
                });
        });
    }

    /**
     * Hiển thị kết quả Process lên giao diện
     * Gồm: danh sách kết quả, cảnh báo, tóm tắt, nút download
     * File lỗi sẽ có nút RETRY để chạy lại
     */
    function displayResults(data) {
        // Danh sách kết quả từng cặp file
        resultsBody.innerHTML = '';
        data.results.forEach(item => {
            const div = document.createElement('div');
            div.className = `result-item ${item.status}`;
            // Nếu lỗi → hiện nút Retry
            const retryBtn = item.status === 'error'
                ? `<button class="btn btn-sm btn-outline-warning btn-retry" data-index="${item.index}" title="Chạy lại file này">
                     <i class="bi bi-arrow-clockwise me-1"></i>Retry
                   </button>`
                : '';
            div.innerHTML = `
                <span class="result-text ${item.status}">
                    <i class="bi ${item.status === 'success' ? 'bi-check-circle-fill' : 'bi-x-circle-fill'} me-2"></i>
                    <strong>Pair ${item.index}:</strong> ${item.message}
                </span>
                <div>
                    ${retryBtn}
                    ${item.zip_file ? `<a href="/api/download/${item.zip_file}" class="btn btn-download btn-sm">
                        <i class="bi bi-download me-1"></i>Download
                    </a>` : ''}
                </div>
            `;
            resultsBody.appendChild(div);
        });
        resultsContainer.style.display = 'block';

        // Gắn sự kiện Retry: scroll đến file lỗi và highlight
        document.querySelectorAll('.btn-retry').forEach(btn => {
            btn.addEventListener('click', function() {
                const idx = this.dataset.index;
                // Scroll đến cặp file tương ứng
                const pair = document.querySelectorAll('.file-pair')[idx - 1];
                if (pair) {
                    pair.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    pair.style.borderColor = '#f77f00';
                    pair.style.boxShadow = '0 0 8px rgba(247,127,0,0.3)';
                    setTimeout(() => {
                        pair.style.borderColor = '';
                        pair.style.boxShadow = '';
                    }, 3000);
                }
                // Thông báo cho user
                showAlert(
                    `Kiểm tra lại file cặp ${idx}, sửa file lỗi rồi nhấn PROCESS lại.`,
                    'info'
                );
            });
        });

        // Cảnh báo NID trùng
        if (data.warnings && data.warnings.length > 0) {
            data.warnings.forEach(w => {
                const wDiv = document.createElement('div');
                wDiv.className = 'warning-item';
                wDiv.innerHTML = `<i class="bi bi-exclamation-triangle-fill me-2"></i>${w}`;
                warningsContainer.appendChild(wDiv);
            });
        }

        // Tóm tắt (total, success, fail, time)
        if (data.summary) {
            document.getElementById('summaryTotal').textContent = data.summary.total;
            document.getElementById('summarySuccess').textContent = data.summary.success;
            document.getElementById('summaryFail').textContent = data.summary.fail;
            document.getElementById('summaryTime').textContent = data.summary.time;
            document.getElementById('summaryPassword').textContent = data.summary.password;
            summaryContainer.style.display = 'block';
        }

        // Nút Download All (khi có nhiều hơn 1 file)
        if (data.zip_files && data.zip_files.length > 1) {
            downloadAllContainer.style.display = 'block';
        }
    }

    /**
     * Hiển thị thông báo alert
     */
    function showAlert(message, type) {
        const alertDiv = document.createElement('div');
        alertDiv.className = `alert alert-${type} alert-dismissible fade show`;
        alertDiv.innerHTML = `
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        `;
        document.querySelector('.page-content').prepend(alertDiv);
        setTimeout(() => alertDiv.remove(), 5000);
    }

    // ============================================================
    // TAB UPDATE - Load danh sách ticket + tạo form update
    // ============================================================

    /** Load danh sách ticket từ server cho autocomplete */
    function loadTickets() {
        fetch('/api/tickets')
            .then(r => r.json())
            .then(data => {
                allTickets = data.tickets;
            });
    }
    loadTickets();

    // Tạo form update khi chọn số lượng
    if (updateCountSelect) {
        updateCountSelect.addEventListener('change', function () {
            const count = parseInt(this.value);
            updatePairs.innerHTML = '';
            if (updateBtn) updateBtn.disabled = (count === 0);

            for (let i = 1; i <= count; i++) {
                const pair = document.createElement('div');
                pair.className = 'file-pair';
                pair.innerHTML = `
                    <div class="file-pair-header">
                        <i class="bi bi-arrow-repeat"></i> Update ${i}
                    </div>
                    <div class="mb-2">
                        <label class="form-label small fw-semibold">Ticket cần update</label>
                        ${createTicketSearchInput(`ticket_name_${i}`, 'Gõ tên ticket để tìm...')}
                    </div>
                    <div class="mb-2">
                        <label class="form-label small fw-semibold">File Result mới</label>
                        <input type="file" class="form-control form-control-sm"
                               name="result_file_${i}" accept=".xlsx,.xls,.csv">
                    </div>
                `;
                updatePairs.appendChild(pair);
            }
            // Gắn autocomplete cho các ô vừa tạo
            initTicketAutocomplete();
        });
    }

    // Preview ticket khi chọn trong ô tìm kiếm bên phải
    if (previewTicketSelect) {
        // Đổi select thành input searchable
        const previewWrapper = previewTicketSelect.parentElement;
        previewWrapper.innerHTML = `
            <div class="position-relative ticket-search-wrapper" style="width:220px;display:inline-block">
                <input type="text" class="form-control form-control-sm ticket-search-input"
                       id="previewTicketInput" placeholder="Gõ ticket để xem..."
                       autocomplete="off">
                <div class="ticket-suggestions" style="display:none;position:absolute;top:100%;left:0;right:0;
                    z-index:1000;max-height:200px;overflow-y:auto;background:#fff;
                    border:1px solid #dee2e6;border-radius:6px;box-shadow:0 4px 12px rgba(0,0,0,0.1);">
                </div>
            </div>
            <small class="text-muted ms-2" id="previewCount">0 records</small>
        `;
        // Gắn autocomplete
        setTimeout(() => {
            initTicketAutocomplete();
            // Khi chọn ticket → load preview
            const previewInput = document.getElementById('previewTicketInput');
            if (previewInput) {
                previewInput.addEventListener('change', function() {
                    loadPreview(this.value.trim());
                });
                // Cũng load khi nhấn Enter
                previewInput.addEventListener('keydown', function(e) {
                    if (e.key === 'Enter') {
                        e.preventDefault();
                        loadPreview(this.value.trim());
                    }
                });
            }
        }, 100);
    }

    function loadPreview(ticket) {
        const pBody = document.getElementById('previewBody');
        const pCount = document.getElementById('previewCount');
        if (!ticket) {
            pBody.innerHTML = '<tr><td colspan="5" class="text-center text-muted py-4">Gõ tên ticket để xem</td></tr>';
            pCount.textContent = '0 records';
            return;
        }
        fetch(`/api/ticket-detail/${encodeURIComponent(ticket)}`)
            .then(r => r.json())
            .then(data => {
                pBody.innerHTML = '';
                pCount.textContent = `${data.total} records`;
                if (data.data.length === 0) {
                    pBody.innerHTML = '<tr><td colspan="5" class="text-center text-muted">Không có dữ liệu</td></tr>';
                    return;
                }
                data.data.forEach(row => {
                    const tr = document.createElement('tr');
                    tr.innerHTML = `
                        <td><strong>${row.NID || ''}</strong></td>
                        <td>${row.NewPhoneNumber || ''}</td>
                        <td>${row.TicketPega || ''}</td>
                        <td><span class="badge ${getResultClass(row.ITResult)}">${row.ITResult || ''}</span></td>
                        <td>${row.DateUpdate || ''}</td>
                    `;
                    pBody.appendChild(tr);
                });
            });
    }

    function getResultClass(result) {
        if (!result) return 'bg-secondary';
        const r = result.toLowerCase();
        if (r === 'success') return 'bg-success';
        if (r === 'failed') return 'bg-danger';
        if (r === 'not found') return 'bg-warning text-dark';
        return 'bg-secondary';
    }

    // ============================================================
    // TAB UPDATE - Nút CẬP NHẬT: validate + gửi lên server
    // ============================================================
    if (updateBtn) {
        updateBtn.addEventListener('click', function () {
            const count = parseInt(updateCountSelect.value);
            if (count === 0) return;

            // Validate
            for (let i = 1; i <= count; i++) {
                const ticketInput = document.querySelector(`input[name="ticket_name_${i}"]`);
                const fileSel = document.querySelector(`input[name="result_file_${i}"]`);
                if (!ticketInput || !ticketInput.value.trim()) {
                    showAlert(`Chưa nhập tên ticket số ${i}`, 'warning');
                    return;
                }
                if (!fileSel || !fileSel.files.length) {
                    showAlert(`Thiếu file Result cho update số ${i}`, 'warning');
                    return;
                }
            }

            // Đóng gói form data
            const formData = new FormData();
            formData.append('update_count', count);
            formData.append('password', document.getElementById('updatePassword').value);
            for (let i = 1; i <= count; i++) {
                const ticketInput = document.querySelector(`input[name="ticket_name_${i}"]`);
                const fileSel = document.querySelector(`input[name="result_file_${i}"]`);
                formData.append(`ticket_name_${i}`, ticketInput.value.trim());
                formData.append(`result_file_${i}`, fileSel.files[0]);
            }

            updateBtn.disabled = true;
            updateBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-2"></span>Đang cập nhật...';

            // Gửi request
            fetch('/api/update-result', { method: 'POST', body: formData })
                .then(r => r.json())
                .then(data => {
                    if (data.success) {
                        // Hiển thị kết quả
                        updateResultsBody.innerHTML = '';
                        data.results.forEach(item => {
                            const div = document.createElement('div');
                            div.className = `result-item ${item.status}`;
                            div.innerHTML = `
                                <span class="result-text ${item.status}">
                                    <i class="bi ${item.status === 'success' ? 'bi-check-circle-fill' : 'bi-x-circle-fill'} me-2"></i>
                                    ${item.message}
                                </span>
                                ${item.zip_file ? `<a href="/api/download/${item.zip_file}" class="btn btn-download btn-sm">
                                    <i class="bi bi-download me-1"></i>Download
                                </a>` : ''}
                            `;
                            updateResultsBody.appendChild(div);
                        });
                        updateResults.style.display = 'block';

                        // Tóm tắt
                        document.getElementById('sumUpdated').textContent = data.summary.total_updated;
                        document.getElementById('sumNotFound').textContent = data.summary.total_not_found;
                        document.getElementById('sumTime').textContent = data.summary.time;
                        updateSummary.style.display = 'block';

                        // Refresh preview
                        if (previewTicketSelect && previewTicketSelect.value) {
                            previewTicketSelect.dispatchEvent(new Event('change'));
                        }
                    } else {
                        showAlert(data.error, 'danger');
                    }
                })
                .catch(err => showAlert('Lỗi: ' + err.message, 'danger'))
                .finally(() => {
                    updateBtn.disabled = false;
                    updateBtn.innerHTML = '<i class="bi bi-arrow-repeat me-2"></i>CẬP NHẬT';
                });
        });
    }
});
