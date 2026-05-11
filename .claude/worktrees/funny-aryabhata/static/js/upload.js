document.addEventListener('DOMContentLoaded', function() {
    const dropArea = document.getElementById('dropArea');
    const fileInput = document.getElementById('fileInput');
    const fileList = document.getElementById('fileList');
    const uploadButton = document.getElementById('uploadButton');
    const uploadForm = document.getElementById('uploadForm');

    let files = [];

    // Prevent default drag behaviors
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, preventDefaults, false);
        document.body.addEventListener(eventName, preventDefaults, false);
    });

    // Highlight drop area when item is dragged over it
    ['dragenter', 'dragover'].forEach(eventName => {
        dropArea.addEventListener(eventName, highlight, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, unhighlight, false);
    });

    // Handle dropped files
    dropArea.addEventListener('drop', handleDrop, false);

    // Handle selected files
    fileInput.addEventListener('change', handleFiles, false);

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    function highlight() {
        dropArea.classList.add('drag-over');
    }

    function unhighlight() {
        dropArea.classList.remove('drag-over');
    }

    function handleDrop(e) {
        const dt = e.dataTransfer;
        const droppedFiles = dt.files;
        fileInput.files = droppedFiles; // Assign files to input element
        handleFiles({ target: fileInput });
    }

    function handleFiles(e) {
        const newFiles = Array.from(e.target.files).filter(file => {
            const fileName = file.name.toLowerCase();
            const isValid = fileName.endsWith('.docx');

            if (!isValid) {
                alert(`File ${file.name} is not a .docx file and will be skipped.`);
                return false;
            }

            // Check for duplicate files
            const isDuplicate = files.some(f =>
                f.name === file.name &&
                f.size === file.size &&
                f.lastModified === file.lastModified
            );
            if (isDuplicate) {
                alert(`File ${file.name} is already in the upload queue.`);
                return false;
            }

            return true;
        });

        files = [...files, ...newFiles];
        updateFileList();
    }

    function updateFileList() {
        fileList.innerHTML = '';

        if (files.length === 0) {
            uploadButton.style.display = 'none';
            return;
        }

        files.forEach((file, index) => {
            const fileItem = document.createElement('div');
            fileItem.className = 'file-item';

            const fileInfo = document.createElement('div');
            fileInfo.className = 'file-info';

            const fileName = document.createElement('span');
            fileName.className = 'file-name';
            fileName.textContent = file.name;

            const fileSize = document.createElement('span');
            fileSize.className = 'file-size';
            fileSize.textContent = formatFileSize(file.size);

            const removeFile = document.createElement('span');
            removeFile.className = 'remove-file';
            removeFile.innerHTML = '<i class="fas fa-times"></i>';
            removeFile.addEventListener('click', (e) => {
                e.preventDefault();
                removeFileFromList(index);
            });

            fileInfo.appendChild(fileName);
            fileInfo.appendChild(fileSize);

            fileItem.appendChild(fileInfo);
            fileItem.appendChild(removeFile);

            fileList.appendChild(fileItem);
        });

        uploadButton.style.display = 'inline-block';
    }

    function removeFileFromList(index) {
        files.splice(index, 1);

        // Update the file input's files property
        const dataTransfer = new DataTransfer();
        files.forEach(file => dataTransfer.items.add(file));
        fileInput.files = dataTransfer.files;

        updateFileList();
    }

    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    // Handle form submission
    uploadForm.addEventListener('submit', function(e) {
        e.preventDefault();

        if (files.length === 0) {
            alert('Please select at least one file to upload.');
            return;
        }

        // Show loading state
        uploadButton.disabled = true;
        uploadButton.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Processing...';

        // Create FormData and append files
        const formData = new FormData();
        files.forEach(file => {
            formData.append('files', file);
        });

        // Add CSRF token if using Flask-WTF
        const csrfToken = document.querySelector('input[name="csrf_token"]')?.value;
        if (csrfToken) {
            formData.append('csrf_token', csrfToken);
        }

        // Submit the form
        fetch(uploadForm.action, {
            method: 'POST',
            body: formData
        })
        .then(response => {
            if (response.redirected) {
                window.location.href = response.url;
            } else {
                return response.json().then(data => {
                    if (data.error) {
                        throw new Error(data.error);
                    }
                    return data;
                });
            }
        })
        .catch(error => {
            console.error('Error:', error);
            alert('Error uploading files: ' + error.message);
        })
        .finally(() => {
            uploadButton.disabled = false;
            uploadButton.innerHTML = '<i class="fas fa-play"></i> Process Files';
        });
    });
});