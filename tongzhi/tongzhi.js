// 通知功能模块

function showNotification(message, type = 'success') {
    const container = document.querySelector('.notification-container');
    const notification = document.createElement('div');
    notification.className = `notification ${type}`;
    notification.textContent = message;
    container.appendChild(notification);
    
    requestAnimationFrame(() => notification.classList.add('show'));
    
    setTimeout(() => {
        notification.classList.remove('show');
        setTimeout(() => notification.remove(), 300);
    }, 3000);
}

// 将函数添加到工具对象中，以便其他模块使用
if (typeof utils !== 'undefined') {
    utils.showNotification = showNotification;
}