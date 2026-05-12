(function () {
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
