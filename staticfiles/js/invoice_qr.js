document.addEventListener('DOMContentLoaded', () => {
    const excelFileInput = document.getElementById('excelFileInput');
    const wordFileInput = document.getElementById('wordFileInput');
    const excelLabel = document.getElementById('excelLabel');
    const wordLabel = document.getElementById('wordLabel');
    const excelError = document.getElementById('excelError');
    const wordError = document.getElementById('wordError');
    const submitBtn = document.getElementById('submitBtn');

    let selectedExcelFile = null;
    let selectedWordFiles = [];

    function isExcelFile(file) {
        const name = (file && file.name || '').toLowerCase();
        return name.endsWith('.xlsx') || name.endsWith('.xls');
    }

    function isWordFile(file) {
        const name = (file && file.name || '').toLowerCase();
        return name.endsWith('.docx');
    }

    function showError(element, message) {
        element.textContent = message;
        element.classList.add('visible');
    }

    function clearError(element) {
        element.textContent = '';
        element.classList.remove('visible');
    }

    function updateFileName(elementId, files) {
        const element = document.getElementById(elementId);

        if (!files || files.length === 0) {
            element.textContent = '';
            return;
        }

        if (files.length === 1) {
            element.textContent = '✓ ' + files[0].name;
        } else {
            element.textContent = '✓ ' + files.length + ' files selected';
        }
    }

    function validateExcelSelection(fileList) {
        if (!fileList || fileList.length === 0) {
            selectedExcelFile = null;
            excelFileInput.value = '';
            clearError(excelError);
            updateFileName('excelFileName', []);
            return null;
        }

        const firstFile = fileList[0];
        if (!isExcelFile(firstFile)) {
            selectedExcelFile = null;
            showError(excelError, '❌ Invalid file type. Please upload an Excel file (.xlsx or .xls).');
            excelFileInput.value = '';
            updateFileName('excelFileName', []);
            return null;
        }

        selectedExcelFile = firstFile;
        clearError(excelError);
        updateFileName('excelFileName', [firstFile]);
        return firstFile;
    }

    function validateWordSelection(fileList) {
        if (!fileList || fileList.length === 0) {
            selectedWordFiles = [];
            wordFileInput.value = '';
            clearError(wordError);
            updateFileName('wordFileName', []);
            return [];
        }

        const validFiles = [];
        for (const file of fileList) {
            if (!isWordFile(file)) {
                selectedWordFiles = validFiles;
                showError(wordError, '❌ Invalid file type. Please upload Word documents (.docx) only.');
                wordFileInput.value = '';
                updateFileName('wordFileName', validFiles);
                return validFiles;
            }
            validFiles.push(file);
        }

        selectedWordFiles = validFiles;
        clearError(wordError);
        updateFileName('wordFileName', validFiles);
        return validFiles;
    }

    excelFileInput.addEventListener('change', function () {
        const selected = this.files ? Array.from(this.files) : [];
        validateExcelSelection(selected);
    });

    wordFileInput.addEventListener('change', function () {
        const selected = this.files ? Array.from(this.files) : [];
        validateWordSelection(selected);
    });

    excelLabel.addEventListener('dragover', (e) => {
        e.preventDefault();
        excelLabel.classList.add('active');
    });

    excelLabel.addEventListener('dragleave', () => {
        excelLabel.classList.remove('active');
    });

    excelLabel.addEventListener('drop', (e) => {
        e.preventDefault();
        excelLabel.classList.remove('active');
        validateExcelSelection(Array.from(e.dataTransfer.files || []));
    });

    wordLabel.addEventListener('dragover', (e) => {
        e.preventDefault();
        wordLabel.classList.add('active');
    });

    wordLabel.addEventListener('dragleave', () => {
        wordLabel.classList.remove('active');
    });

    wordLabel.addEventListener('drop', (e) => {
        e.preventDefault();
        wordLabel.classList.remove('active');
        validateWordSelection(Array.from(e.dataTransfer.files || []));
    });

    async function uploadFiles() {
        if (!selectedExcelFile) {
            const status = document.getElementById('status');
            status.innerText = '❌ Please upload an Excel file first.';
            status.classList.add('error');
            return;
        }

        if (selectedWordFiles.length === 0) {
            const status = document.getElementById('status');
            status.innerText = '❌ Please upload at least one Word document.';
            status.classList.add('error');
            return;
        }

        const formData = new FormData();
        formData.append('data_file', selectedExcelFile);
        selectedWordFiles.forEach(file => formData.append('templates', file));

        console.log('Excel file:', selectedExcelFile);
        console.log('Word files:', selectedWordFiles);
        console.log('Word file count:', selectedWordFiles.length);
        console.log('FormData contents:');
        for (const [key, value] of formData.entries()) {
            console.log(key, value instanceof File ? `${value.name} (${value.type}, ${value.size} bytes)` : value);
        }

        const progressContainer = document.getElementById('progressContainer');
        const progress = document.getElementById('progress');
        const statusElement = document.getElementById('status');
        const csrfToken = document.querySelector('input[name="csrfmiddlewaretoken"]')?.value || '';

        progressContainer.style.display = 'block';
        progress.style.width = '10%';
        submitBtn.disabled = true;
        statusElement.innerText = '⏳ Processing your files...';
        statusElement.classList.remove('success', 'error');

        try {
            const res = await fetch('/upload/', {
                method: 'POST',
                body: formData,
                headers: {
                    'X-CSRFToken': csrfToken
                }
            });

            progress.style.width = '70%';
            const data = await res.json();

            if (!res.ok || data.success === false) {
                statusElement.innerText = data.error || '❌ Invalid file type.';
                statusElement.classList.add('error');
                progressContainer.style.display = 'none';
                submitBtn.disabled = false;
                return;
            }

            progress.style.width = '100%';
            setTimeout(() => {
                progressContainer.style.display = 'none';
                submitBtn.disabled = false;
            }, 800);

            statusElement.innerText = '✅ Processing complete!';
            statusElement.classList.add('success');

            const container = document.getElementById('results');
            container.innerHTML = '';
            document.getElementById('downloadAllContainer').style.display = 'block';

            data.results.forEach(item => {
                if (item.status === 'Success') {
                    const downloadUrl = item.download_id ? `/download/${item.download_id}/` : '#';
                    container.innerHTML += `
                        <div class="result-card">
                            <b class="success-icon">${item.file}</b>
                            <div class="btn-group">
                                <a class="btn-download" href="${downloadUrl}" download>📥 Download Document</a>
                            </div>
                        </div>
                    `;
                } else {
                    container.innerHTML += `
                        <div class="result-card error">
                            <b class="error-icon">${item.file}</b>
                            <p style="margin-top:10px;color:#dc3545;">${item.status}</p>
                        </div>
                    `;
                }
            });
        } catch (error) {
            console.error(error);
            statusElement.innerText = '❌ An error occurred while processing files';
            statusElement.classList.add('error');
            submitBtn.disabled = false;
            progressContainer.style.display = 'none';
        }
    }

    submitBtn.addEventListener('click', uploadFiles);

    document.getElementById('clearAllBtn')?.addEventListener('click', async function () {
        const statusElement = document.getElementById('status');
        const progressContainer = document.getElementById('progressContainer');
        const progress = document.getElementById('progress');
        const csrfToken = document.querySelector('input[name="csrfmiddlewaretoken"]')?.value || '';

        progressContainer.style.display = 'block';
        progress.style.width = '20%';
        statusElement.innerText = '⏳ Clearing current batch...';
        statusElement.classList.remove('success', 'error');

        try {
            const res = await fetch('/clear-all/', {
                method: 'POST',
                headers: {
                    'X-CSRFToken': csrfToken,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({})
            });

            const data = await res.json();
            progress.style.width = '100%';

            setTimeout(() => {
                document.getElementById('results').innerHTML = '';
                document.getElementById('downloadAllContainer').style.display = 'none';
                selectedExcelFile = null;
                selectedWordFiles = [];
                excelFileInput.value = '';
                wordFileInput.value = '';
                clearError(excelError);
                clearError(wordError);
                updateFileName('excelFileName', []);
                updateFileName('wordFileName', []);
                document.getElementById('progress').style.width = '0%';
                progressContainer.style.display = 'none';
                submitBtn.disabled = false;
            }, 300);

            statusElement.innerText = data.message || 'Current batch cleared.';
            statusElement.classList.add('success');
        } catch (error) {
            statusElement.innerText = '❌ Unable to clear the current batch.';
            statusElement.classList.add('error');
            progressContainer.style.display = 'none';
        }
    });
});
