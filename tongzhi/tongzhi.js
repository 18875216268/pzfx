// 通知模块
const notification = {
    show(message, type = 'success') {
        const container = document.querySelector('.notification-container');
        const notification = document.createElement('div');
        notification.className = `notification ${type}`;
        notification.textContent = message;
        container.appendChild(notification);
        
        // 使用 requestAnimationFrame 确保动画流畅
        requestAnimationFrame(() => notification.classList.add('show'));
        
        // 3秒后自动关闭
        setTimeout(() => {
            notification.classList.remove('show');
            setTimeout(() => notification.remove(), 300);
        }, 3000);
    }
};