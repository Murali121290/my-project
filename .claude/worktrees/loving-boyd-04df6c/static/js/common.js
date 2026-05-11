// Common JavaScript functionality for S4carlisle application

document.addEventListener('DOMContentLoaded', function() {
    // Initialize all common functionality
    initFlashMessages();
    initBubblesAnimation();
    initNavigation();
    initFormValidation();
    initAnimations();
});

// Flash Messages functionality
function initFlashMessages() {
    const closeButtons = document.querySelectorAll('.flash .close');
    
    closeButtons.forEach(button => {
        button.addEventListener('click', function() {
            const flash = this.closest('.flash');
            if (flash) {
                flash.style.opacity = '0';
                flash.style.transform = 'translateY(-10px)';
                setTimeout(() => flash.remove(), 300);
            }
        });
    });
    
    // Auto-remove flash messages after 5 seconds
    const flashMessages = document.querySelectorAll('.flash');
    flashMessages.forEach(flash => {
        setTimeout(() => {
            if (flash.parentElement) {
                flash.style.opacity = '0';
                flash.style.transform = 'translateY(-10px)';
                setTimeout(() => flash.remove(), 300);
            }
        }, 5000);
    });
}

// Bubble animation for background effects
function initBubblesAnimation() {
    const bubblesContainers = document.querySelectorAll('.bubbles');
    
    bubblesContainers.forEach(container => {
        // Clear any existing bubbles
        container.innerHTML = '';
        
        // Create bubbles based on container size
        const bubbleCount = Math.floor(container.clientWidth / 50);
        
        for (let i = 0; i < bubbleCount; i++) {
            const bubble = document.createElement('div');
            bubble.classList.add('bubble');
            
            // Random properties for variety
            const size = Math.random() * 80 + 20;
            const posX = Math.random() * 100;
            const delay = Math.random() * 15;
            const duration = 10 + Math.random() * 20;
            const opacity = Math.random() * 0.2 + 0.1;
            
            bubble.style.width = `${size}px`;
            bubble.style.height = `${size}px`;
            bubble.style.left = `${posX}%`;
            bubble.style.animationDelay = `${delay}s`;
            bubble.style.animationDuration = `${duration}s`;
            bubble.style.opacity = opacity;
            
            container.appendChild(bubble);
        }
    });
}

// Navigation functionality
function initNavigation() {
    // Mobile navigation toggle
    const navToggle = document.getElementById('navToggle');
    const navMenu = document.getElementById('navMenu');
    
    if (navToggle && navMenu) {
        navToggle.addEventListener('click', function() {
            navMenu.classList.toggle('active');
            navToggle.classList.toggle('active');
        });
    }
    
    // Active navigation link highlighting
    const currentPath = window.location.pathname;
    const navLinks = document.querySelectorAll('.nav-link');
    
    navLinks.forEach(link => {
        const href = link.getAttribute('href');
        if (href && currentPath.includes(href) && href !== '/') {
            link.classList.add('active');
        }
    });
}

// Form validation helpers
function initFormValidation() {
    const forms = document.querySelectorAll('form');
    
    forms.forEach(form => {
        form.addEventListener('submit', function(e) {
            const requiredFields = form.querySelectorAll('[required]');
            let isValid = true;
            
            requiredFields.forEach(field => {
                if (!field.value.trim()) {
                    isValid = false;
                    highlightFieldError(field);
                } else {
                    clearFieldError(field);
                }
            });
            
            // Password confirmation validation
            const password = form.querySelector('input[type="password"][name="password"]');
            const confirmPassword = form.querySelector('input[type="password"][name="confirm_password"]');
            
            if (password && confirmPassword && password.value !== confirmPassword.value) {
                isValid = false;
                highlightFieldError(confirmPassword, 'Passwords do not match');
            }
            
            if (!isValid) {
                e.preventDefault();
                showAlert('Please fill in all required fields correctly.', 'error');
            }
        });
    });
    
    // Real-time validation
    const inputs = document.querySelectorAll('input, textarea, select');
    inputs.forEach(input => {
        input.addEventListener('blur', function() {
            if (this.hasAttribute('required') && !this.value.trim()) {
                highlightFieldError(this);
            } else {
                clearFieldError(this);
            }
        });
        
        input.addEventListener('input', function() {
            clearFieldError(this);
        });
    });
}

function highlightFieldError(field, message = 'This field is required') {
    field.style.borderColor = 'var(--danger)';
    
    // Remove existing error message
    const existingError = field.parentElement.querySelector('.field-error');
    if (existingError) {
        existingError.remove();
    }
    
    // Add error message
    const errorDiv = document.createElement('div');
    errorDiv.className = 'field-error';
    errorDiv.style.color = 'var(--danger)';
    errorDiv.style.fontSize = '0.875rem';
    errorDiv.style.marginTop = '0.25rem';
    errorDiv.textContent = message;
    
    field.parentElement.appendChild(errorDiv);
}

function clearFieldError(field) {
    field.style.borderColor = '';
    
    const errorDiv = field.parentElement.querySelector('.field-error');
    if (errorDiv) {
        errorDiv.remove();
    }
}

// Animation triggers
function initAnimations() {
    // Intersection Observer for scroll animations
    const animatedElements = document.querySelectorAll('.fade-in, .slide-in-left, .slide-in-right');
    
    if ('IntersectionObserver' in window) {
        const observer = new IntersectionObserver((entries) => {
            entries.forEach(entry => {
                if (entry.isIntersecting) {
                    entry.target.style.opacity = '1';
                    entry.target.style.transform = 'translate(0, 0)';
                    observer.unobserve(entry.target);
                }
            });
        }, { threshold: 0.1 });
        
        animatedElements.forEach(el => {
            observer.observe(el);
        });
    } else {
        // Fallback for browsers without Intersection Observer
        animatedElements.forEach(el => {
            el.style.opacity = '1';
            el.style.transform = 'translate(0, 0)';
        });
    }
}

// Alert/Notification system
function showAlert(message, type = 'info', duration = 5000) {
    const alertDiv = document.createElement('div');
    alertDiv.className = `flash ${type}`;
    alertDiv.innerHTML = `
        ${message}
        <button class="close">&times;</button>
    `;
    
    const flashContainer = document.querySelector('.flash-messages') || createFlashContainer();
    flashContainer.appendChild(alertDiv);
    
    // Initialize close functionality
    const closeBtn = alertDiv.querySelector('.close');
    closeBtn.addEventListener('click', () => removeAlert(alertDiv));
    
    // Auto-remove after duration
    if (duration > 0) {
        setTimeout(() => removeAlert(alertDiv), duration);
    }
    
    return alertDiv;
}

function removeAlert(alertDiv) {
    alertDiv.style.opacity = '0';
    alertDiv.style.transform = 'translateY(-10px)';
    setTimeout(() => {
        if (alertDiv.parentElement) {
            alertDiv.remove();
        }
    }, 300);
}

function createFlashContainer() {
    const container = document.createElement('div');
    container.className = 'flash-messages';
    document.querySelector('.main-content').prepend(container);
    return container;
}

// File handling utilities
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function validateFileType(file, allowedTypes) {
    const extension = file.name.toLowerCase().split('.').pop();
    return allowedTypes.includes('.' + extension);
}

// API helpers
async function apiRequest(url, options = {}) {
    try {
        const response = await fetch(url, {
            headers: {
                'Content-Type': 'application/json',
                'X-Requested-With': 'XMLHttpRequest',
                ...options.headers
            },
            ...options
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const data = await response.json();
        return data;
    } catch (error) {
        console.error('API request failed:', error);
        showAlert('An error occurred while processing your request.', 'error');
        throw error;
    }
}

// Local storage utilities
function setStorage(key, value) {
    try {
        localStorage.setItem(key, JSON.stringify(value));
        return true;
    } catch (error) {
        console.error('Error saving to localStorage:', error);
        return false;
    }
}

function getStorage(key, defaultValue = null) {
    try {
        const item = localStorage.getItem(key);
        return item ? JSON.parse(item) : defaultValue;
    } catch (error) {
        console.error('Error reading from localStorage:', error);
        return defaultValue;
    }
}

function removeStorage(key) {
    try {
        localStorage.removeItem(key);
        return true;
    } catch (error) {
        console.error('Error removing from localStorage:', error);
        return false;
    }
}

// Debounce function for performance
function debounce(func, wait, immediate) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            timeout = null;
            if (!immediate) func.apply(this, args);
        };
        const callNow = immediate && !timeout;
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
        if (callNow) func.apply(this, args);
    };
}

// Throttle function for performance
function throttle(func, limit) {
    let inThrottle;
    return function(...args) {
        if (!inThrottle) {
            func.apply(this, args);
            inThrottle = true;
            setTimeout(() => inThrottle = false, limit);
        }
    };
}

// Export functions for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        showAlert,
        removeAlert,
        formatFileSize,
        validateFileType,
        apiRequest,
        setStorage,
        getStorage,
        removeStorage,
        debounce,
        throttle
    };
}