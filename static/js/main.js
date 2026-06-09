document.addEventListener('DOMContentLoaded', () => {
    // 1. Mobile Sidebar Toggle
    const sidebar = document.querySelector('.sidebar');
    const sidebarToggle = document.querySelector('.sidebar-toggle');
    
    if (sidebarToggle && sidebar) {
        // Create overlay element
        const overlay = document.createElement('div');
        overlay.className = 'sidebar-overlay';
        overlay.style.cssText = `
            position: fixed;
            top: 0; left: 0; right: 0; bottom: 0;
            background-color: rgba(15, 23, 42, 0.4);
            backdrop-filter: blur(2px);
            z-index: 95;
            opacity: 0;
            pointer-events: none;
            transition: opacity 0.3s ease;
        `;
        document.body.appendChild(overlay);
        
        sidebarToggle.addEventListener('click', () => {
            sidebar.classList.toggle('open');
            if (sidebar.classList.contains('open')) {
                overlay.style.opacity = '1';
                overlay.style.pointerEvents = 'auto';
            } else {
                overlay.style.opacity = '0';
                overlay.style.pointerEvents = 'none';
            }
        });
        
        overlay.addEventListener('click', () => {
            sidebar.classList.remove('open');
            overlay.style.opacity = '0';
            overlay.style.pointerEvents = 'none';
        });
    }
    
    // 2. Alert Banner Auto-dismiss
    const alerts = document.querySelectorAll('.alert-banner');
    alerts.forEach(alert => {
        setTimeout(() => {
            alert.style.transition = 'all 0.5s ease';
            alert.style.opacity = '0';
            alert.style.transform = 'translateY(-20px)';
            alert.style.height = '0';
            alert.style.padding = '0';
            alert.style.marginBottom = '0';
            alert.style.border = 'none';
            setTimeout(() => {
                alert.remove();
            }, 500);
        }, 4000);
    });

    // 3. General Search Filter Utility (Supports tables and card grids)
    const searchInputs = document.querySelectorAll('[data-search-table]');
    searchInputs.forEach(input => {
        const targetId = input.getAttribute('data-search-table');
        const container = document.getElementById(targetId);
        if (!container) return;
        
        input.addEventListener('input', (e) => {
            const query = e.target.value.toLowerCase().trim();
            
            if (container.tagName.toLowerCase() === 'table') {
                const rows = container.querySelectorAll('tbody tr');
                rows.forEach(row => {
                    const text = row.textContent.toLowerCase();
                    row.style.display = text.includes(query) ? '' : 'none';
                });
            } else {
                // Search grid cards
                const items = container.querySelectorAll('.stock-card, .card');
                items.forEach(item => {
                    const text = item.textContent.toLowerCase();
                    item.style.display = text.includes(query) ? '' : 'none';
                });
            }
        });
    });
});

// 4. Modal Functions (Global Scope)
function openModal(modalId) {
    const modalBackdrop = document.getElementById(modalId);
    if (modalBackdrop) {
        modalBackdrop.classList.add('open');
        document.body.style.overflow = 'hidden'; // Prevent scrolling
    }
}

function closeModal(modalId) {
    const modalBackdrop = document.getElementById(modalId);
    if (modalBackdrop) {
        modalBackdrop.classList.remove('open');
        document.body.style.overflow = ''; // Re-enable scrolling
    }
}
