(() => {
  const loginForm = document.getElementById('loginForm');
  const passwordInput = document.getElementById('passwordInput');
  const togglePasswordBtn = document.getElementById('togglePasswordBtn');
  const loginSubmitBtn = document.getElementById('loginSubmitBtn');
  const loginError = document.getElementById('loginError');
  const loginErrorMessage = document.getElementById('loginErrorMessage');

  function showError(message) {
    loginErrorMessage.textContent = message;
    loginError.classList.remove('hidden');
  }

  function hideError() {
    loginError.classList.add('hidden');
    loginErrorMessage.textContent = '';
  }

  togglePasswordBtn.addEventListener('click', () => {
    const isHidden = passwordInput.type === 'password';
    passwordInput.type = isHidden ? 'text' : 'password';
    togglePasswordBtn.textContent = isHidden ? 'Hide' : 'Show';
    togglePasswordBtn.setAttribute('aria-label', isHidden ? 'Hide password' : 'Show password');
    togglePasswordBtn.setAttribute('aria-pressed', String(isHidden));
    passwordInput.focus();
  });

  loginForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    hideError();

    const password = passwordInput.value;
    if (!password) {
      showError('Please enter your password.');
      passwordInput.focus();
      return;
    }

    loginSubmitBtn.disabled = true;
    loginSubmitBtn.textContent = 'Signing in...';

    try {
      const response = await fetch('/api/auth/login', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ password }),
      });

      const payload = await response.json().catch(() => ({}));
      if (!response.ok) {
        throw new Error(payload.error || 'Invalid password.');
      }

      window.location.assign('/');
    } catch (err) {
      showError(err.message || 'Unable to sign in.');
      loginSubmitBtn.disabled = false;
      loginSubmitBtn.textContent = 'Unlock App';
      passwordInput.focus();
      passwordInput.select();
    }
  });
})();

