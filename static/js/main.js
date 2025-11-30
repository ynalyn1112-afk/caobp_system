// Main JavaScript for CAOBP System

document.addEventListener('DOMContentLoaded', function() {
    // Initialize all components
    initializeToasts();
    initializeAnimations();
    initializeForms();
    initializeThemeToggle();
});

// Toast Notifications
function initializeToasts() {
    // Auto-hide toasts after 5 seconds
    setTimeout(() => {
        const toasts = document.querySelectorAll('.toast');
        toasts.forEach(toast => {
            const bsToast = new bootstrap.Toast(toast);
            bsToast.hide();
        });
    }, 5000);
}

// Show custom toast
function showToast(message, type = 'info', duration = 5000) {
    const toastContainer = document.getElementById('toast-container');
    if (!toastContainer) return;
    
    const toastId = 'toast-' + Date.now();
    const iconClass = getToastIcon(type);
    const headerClass = getToastHeaderClass(type);
    
    const toastHtml = `
        <div class="toast show" id="${toastId}" role="alert" aria-live="assertive" aria-atomic="true">
            <div class="toast-header ${headerClass}">
                <i class="bi ${iconClass} me-2"></i>
                <strong class="me-auto">${getToastTitle(type)}</strong>
                <button type="button" class="btn-close" data-bs-dismiss="toast"></button>
            </div>
            <div class="toast-body">
                ${message}
            </div>
        </div>
    `;
    
    toastContainer.insertAdjacentHTML('beforeend', toastHtml);
    
    // Auto-hide after specified duration
    setTimeout(() => {
        const toast = document.getElementById(toastId);
        if (toast) {
            const bsToast = new bootstrap.Toast(toast);
            bsToast.hide();
        }
    }, duration);
}

function getToastIcon(type) {
    const icons = {
        'success': 'bi-check-circle-fill text-success',
        'error': 'bi-exclamation-triangle-fill text-danger',
        'warning': 'bi-exclamation-triangle-fill text-warning',
        'info': 'bi-info-circle-fill text-info'
    };
    return icons[type] || icons['info'];
}

function getToastHeaderClass(type) {
    const classes = {
        'success': 'bg-success text-white',
        'error': 'bg-danger text-white',
        'warning': 'bg-warning text-dark',
        'info': 'bg-info text-white'
    };
    return classes[type] || classes['info'];
}

function getToastTitle(type) {
    const titles = {
        'success': 'Success',
        'error': 'Error',
        'warning': 'Warning',
        'info': 'Info'
    };
    return titles[type] || 'Info';
}

// Animation System
function initializeAnimations() {
    // Intersection Observer for fade-in animations
    const observerOptions = {
        threshold: 0.1,
        rootMargin: '0px 0px -50px 0px'
    };

    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.classList.add('visible');
            }
        });
    }, observerOptions);

    // Observe all elements with animation classes
    document.querySelectorAll('.fade-in, .slide-in-left, .slide-in-right').forEach(element => {
        observer.observe(element);
    });
}

// Form Enhancements
function initializeForms() {
    // Add loading states to forms
    document.querySelectorAll('form').forEach(form => {
        form.addEventListener('submit', function(e) {
            const submitBtn = form.querySelector('button[type="submit"]');
            if (submitBtn && !submitBtn.disabled) {
                addLoadingState(submitBtn);
            }
        });
    });
    
    // Real-time form validation
    document.querySelectorAll('input[required]').forEach(input => {
        input.addEventListener('blur', validateField);
        input.addEventListener('input', function(e) {
            clearFieldError(e.target);
        });
    });
}

function addLoadingState(button) {
    const originalText = button.innerHTML;
    const loadingSpinner = '<div class="spinner-border spinner-border-sm me-2" role="status"><span class="visually-hidden">Loading...</span></div>';
    
    button.disabled = true;
    button.innerHTML = loadingSpinner + 'Processing...';
    
    // Store original text for potential restoration
    button.setAttribute('data-original-text', originalText);
}

function removeLoadingState(button) {
    const originalText = button.getAttribute('data-original-text');
    if (originalText) {
        button.innerHTML = originalText;
        button.removeAttribute('data-original-text');
    }
    button.disabled = false;
}

function validateField(e) {
    const field = e.target;
    const value = field.value.trim();
    const isValid = field.checkValidity();
    
    if (!isValid) {
        showFieldError(field, 'This field is required');
    } else {
        clearFieldError(field);
    }
}

function showFieldError(field, message) {
    clearFieldError(field);
    
    const errorDiv = document.createElement('div');
    errorDiv.className = 'invalid-feedback d-block';
    errorDiv.textContent = message;
    
    field.classList.add('is-invalid');
    field.parentNode.appendChild(errorDiv);
}

function clearFieldError(field) {
    try {
        if (!field || typeof field !== 'object' || !field.parentNode) {
            console.warn('clearFieldError: Invalid field parameter');
            return;
        }
        
        if (field.classList && typeof field.classList.remove === 'function') {
            field.classList.remove('is-invalid');
        }
        
        const errorDiv = field.parentNode.querySelector('.invalid-feedback');
        if (errorDiv && typeof errorDiv.remove === 'function') {
            errorDiv.remove();
        }
    } catch (error) {
        console.warn('Error in clearFieldError:', error);
    }
}

// Theme Toggle
function initializeThemeToggle() {
    const themeToggle = document.getElementById('themeToggle');
    if (themeToggle) {
        themeToggle.addEventListener('click', toggleTheme);
        
        // Load saved theme
        const savedTheme = localStorage.getItem('theme') || 'light';
        document.body.setAttribute('data-theme', savedTheme);
        updateThemeIcon(savedTheme);
    }
}

function toggleTheme() {
    const currentTheme = document.body.getAttribute('data-theme');
    const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
    
    document.body.setAttribute('data-theme', newTheme);
    localStorage.setItem('theme', newTheme);
    updateThemeIcon(newTheme);
    
    // Show toast notification
    showToast(`Switched to ${newTheme} theme`, 'info', 2000);
}

function updateThemeIcon(theme) {
    const themeToggle = document.getElementById('themeToggle');
    if (themeToggle) {
        const icon = themeToggle.querySelector('i');
        if (icon) {
            icon.className = theme === 'dark' ? 'bi bi-sun' : 'bi bi-moon';
        }
    }
}

// Utility Functions
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

function throttle(func, limit) {
    let inThrottle;
    return function() {
        const args = arguments;
        const context = this;
        if (!inThrottle) {
            func.apply(context, args);
            inThrottle = true;
            setTimeout(() => inThrottle = false, limit);
        }
    };
}

// AJAX Helper
function makeAjaxRequest(url, options = {}) {
    const defaultOptions = {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'X-CSRFToken': getCSRFToken()
        }
    };
    
    const mergedOptions = { ...defaultOptions, ...options };
    
    return fetch(url, mergedOptions)
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.json();
        })
        .catch(error => {
            console.error('AJAX Error:', error);
            showToast('An error occurred. Please try again.', 'error');
            throw error;
        });
}

function getCSRFToken() {
    const token = document.querySelector('[name=csrfmiddlewaretoken]');
    return token ? token.value : '';
}

// Password Strength Checker
function checkPasswordStrength(password) {
    let score = 0;
    let feedback = [];
    
    if (password.length >= 8) score++;
    else feedback.push('At least 8 characters');
    
    if (/[A-Z]/.test(password)) score++;
    else feedback.push('Uppercase letter');
    
    if (/[a-z]/.test(password)) score++;
    else feedback.push('Lowercase letter');
    
    if (/\d/.test(password)) score++;
    else feedback.push('Number');
    
    if (/[!@#$%^&*(),.?":{}|<>]/.test(password)) score++;
    else feedback.push('Special character');
    
    const strength = {
        score: score,
        level: score < 2 ? 'weak' : score < 4 ? 'fair' : score < 5 ? 'good' : 'strong',
        feedback: feedback
    };
    
    return strength;
}

// Smooth Scrolling
function smoothScrollTo(target) {
    const element = document.querySelector(target);
    if (element) {
        element.scrollIntoView({
            behavior: 'smooth',
            block: 'start'
        });
    }
}

// Initialize smooth scrolling for anchor links
document.addEventListener('click', function(e) {
    if (e.target.matches('a[href^="#"]')) {
        e.preventDefault();
        const target = e.target.getAttribute('href');
        smoothScrollTo(target);
    }
});

// Export functions for global use
window.CAOBP = {
    showToast,
    addLoadingState,
    removeLoadingState,
    makeAjaxRequest,
    checkPasswordStrength,
    smoothScrollTo
};
