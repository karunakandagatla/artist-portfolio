function initSiteNav() {
  const menuBtn = document.getElementById('menu-btn');
  const mobileMenu = document.getElementById('mobile-menu');
  if (menuBtn && mobileMenu) {
    menuBtn.addEventListener('click', () => mobileMenu.classList.toggle('hidden'));
  }
}

function closeMobile() {
  const mobileMenu = document.getElementById('mobile-menu');
  if (mobileMenu) mobileMenu.classList.add('hidden');
}

document.addEventListener('DOMContentLoaded', () => {
  const year = document.getElementById('year');
  if (year) year.textContent = new Date().getFullYear();
  initSiteNav();
});
