/**
 * AUHSE — Main JavaScript
 * Mobile menu, user dropdown, smooth scroll, toast auto-dismiss
 */

(function () {
  'use strict';

  // ── Mobile Menu Toggle ──────────────────────────────────────
  var toggle = document.getElementById('navbar-toggle');
  var nav = document.getElementById('navbar-nav');

  if (toggle && nav) {
    toggle.addEventListener('click', function () {
      nav.classList.toggle('open');
      toggle.classList.toggle('active');
    });

    // Close on link click
    nav.querySelectorAll('a.nav-link, a.btn').forEach(function (link) {
      link.addEventListener('click', function () {
        nav.classList.remove('open');
        toggle.classList.remove('active');
      });
    });
  }

  // ── User Dropdown ───────────────────────────────────────────
  var userBtn = document.getElementById('user-menu-btn');
  var userDropdown = document.getElementById('user-dropdown');

  if (userBtn && userDropdown) {
    userBtn.addEventListener('click', function (e) {
      e.stopPropagation();
      userDropdown.classList.toggle('open');
    });

    document.addEventListener('click', function () {
      userDropdown.classList.remove('open');
    });
  }

  // ── Smooth Scroll for Anchor Links ──────────────────────────
  document.querySelectorAll('a[href^="#"]').forEach(function (link) {
    link.addEventListener('click', function (e) {
      var target = document.querySelector(this.getAttribute('href'));
      if (target) {
        e.preventDefault();
        target.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    });
  });

  // ── Sticky Navbar Shadow ────────────────────────────────────
  var navbar = document.getElementById('navbar');
  if (navbar) {
    window.addEventListener('scroll', function () {
      if (window.scrollY > 10) {
        navbar.style.boxShadow = '0 2px 12px rgba(0,0,0,.4)';
      } else {
        navbar.style.boxShadow = 'none';
      }
    });
  }

  // ── Toast Auto-Dismiss ──────────────────────────────────────
  var toasts = document.querySelectorAll('.toast');
  toasts.forEach(function (toast) {
    setTimeout(function () {
      toast.style.transition = 'opacity 0.3s, transform 0.3s';
      toast.style.opacity = '0';
      toast.style.transform = 'translateX(20px)';
      setTimeout(function () { toast.remove(); }, 300);
    }, 6000);
  });

})();
