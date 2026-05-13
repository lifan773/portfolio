(function () {
  /* ====== Mobile Menu ====== */
  var hamburger = document.getElementById('hamburger');
  var navLinks = document.getElementById('navLinks');

  if (hamburger && navLinks) {
    hamburger.addEventListener('click', function () {
      hamburger.classList.toggle('open');
      navLinks.classList.toggle('open');
      document.body.style.overflow = navLinks.classList.contains('open') ? 'hidden' : '';
    });

    // Close menu when a link is clicked
    navLinks.querySelectorAll('a').forEach(function (link) {
      link.addEventListener('click', function () {
        hamburger.classList.remove('open');
        navLinks.classList.remove('open');
        document.body.style.overflow = '';
      });
    });
  }

  /* ====== Nav scroll shadow ====== */
  var navbar = document.getElementById('navbar');
  if (navbar) {
    window.addEventListener('scroll', function () {
      if (window.scrollY > 40) {
        navbar.classList.add('scrolled');
      } else {
        navbar.classList.remove('scrolled');
      }
    });
  }

  /* ====== Scroll Reveal ====== */
  var revealEls = document.querySelectorAll('.reveal');
  var staggerContainers = document.querySelectorAll('.reveal-stagger');

  function revealOnScroll() {
    var windowHeight = window.innerHeight;
    var threshold = 0.12;

    revealEls.forEach(function (el) {
      var top = el.getBoundingClientRect().top;
      if (top < windowHeight * (1 - threshold)) {
        el.classList.add('visible');
      }
    });

    staggerContainers.forEach(function (container) {
      var children = container.children;
      for (var i = 0; i < children.length; i++) {
        var top = children[i].getBoundingClientRect().top;
        if (top < windowHeight * (1 - threshold)) {
          children[i].style.transitionDelay = i * 0.12 + 's';
          children[i].classList.add('visible');
        }
      }
    });
  }

  // Fire on load and scroll
  revealOnScroll();
  window.addEventListener('scroll', revealOnScroll, { passive: true });

  /* ====== Contact form ====== */
  var form = document.getElementById('contactForm');
  if (!form) return;

  var msg = document.getElementById('formMessage');

  form.addEventListener('submit', function (e) {
    e.preventDefault();

    msg.className = 'form-message';
    msg.textContent = '';

    var name = form.name.value.trim();
    var email = form.email.value.trim();
    var body = form.message.value.trim();

    if (!name || !email || !body) {
      msg.className = 'form-message error';
      msg.textContent = 'Please fill in all fields.';
      return;
    }

    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      msg.className = 'form-message error';
      msg.textContent = 'Please enter a valid email address.';
      return;
    }

    msg.className = 'form-message success';
    msg.textContent = 'Thanks, ' + name + '! Your message has been sent.';
    form.reset();
  });
})();
