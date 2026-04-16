// Mobile navigation toggle
const menuBtn = document.getElementById('menu-btn');
const mobileMenu = document.getElementById('mobile-menu');

if (menuBtn && mobileMenu) {
  menuBtn.addEventListener('click', () => {
    mobileMenu.classList.toggle('hidden');
  });
}

// Highlight active nav link based on current page
const currentPath = window.location.pathname.split('/').pop() || 'index.html';
document.querySelectorAll('nav a[data-page]').forEach(link => {
  if (link.dataset.page === currentPath) {
    link.classList.add('text-orange', 'border-b-2', 'border-orange');
  }
});
