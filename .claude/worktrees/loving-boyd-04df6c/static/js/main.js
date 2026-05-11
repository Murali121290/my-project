// Create floating bubbles
function createBubbles(container, count = 15) {
    for (let i = 0; i < count; i++) {
        const bubble = document.createElement('div');
        bubble.classList.add('bubble');

        // Random properties
        const size = Math.random() * 100 + 50;
        const posX = Math.random() * 100;
        const delay = Math.random() * 15;
        const duration = 10 + Math.random() * 20;

        bubble.style.width = `${size}px`;
        bubble.style.height = `${size}px`;
        bubble.style.left = `${posX}%`;
        bubble.style.animationDelay = `${delay}s`;
        bubble.style.animationDuration = `${duration}s`;

        container.appendChild(bubble);
    }
}

// Initialize bubbles on pages that need them
document.addEventListener('DOMContentLoaded', function() {
    const bubblesContainer = document.querySelector('.bubbles');
    if (bubblesContainer) {
        createBubbles(bubblesContainer);
    }
    
    // Close flash messages
    document.querySelectorAll('.flash .close').forEach(button => {
        button.addEventListener('click', function() {
            this.parentElement.style.opacity = '0';
            setTimeout(() => this.parentElement.remove(), 300);
        });
    });

    // Auto-hide flash messages after 5 seconds
    setTimeout(() => {
        document.querySelectorAll('.flash').forEach(flash => {
            flash.style.opacity = '0';
            setTimeout(() => flash.remove(), 300);
        });
    }, 5000);
});

// Format file size
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// Show alert message
function showAlert(message, type = 'info') {
    const alertDiv = document.createElement('div');
    alertDiv.className = `flash ${type}`;
    alertDiv.innerHTML = `
        ${message}
        <button class="close">&times;</button>
    `;

    const flashContainer = document.querySelector('.flash-messages') || document.createElement('div');
    if (!document.querySelector('.flash-messages')) {
        flashContainer.className = 'flash-messages';
        document.querySelector('.main-content').prepend(flashContainer);
    }

    flashContainer.appendChild(alertDiv);

    alertDiv.querySelector('.close').addEventListener('click', () => {
        alertDiv.remove();
    });

    setTimeout(() => alertDiv.remove(), 5000);
}