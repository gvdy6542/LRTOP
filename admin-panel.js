const loginBtn = document.getElementById('loginBtn');
const loginError = document.getElementById('loginError');
const loginContainer = document.querySelector('.login-container');
const modal = document.getElementById('adminModal');
const closeModal = document.getElementById('closeModal');
const buttonsArea = document.getElementById('buttonsArea');
const addBtn = document.getElementById('addBtn');
const newBtnText = document.getElementById('newBtnText');

loginBtn.addEventListener('click', async () => {
  const username = document.getElementById('username').value.trim();
  const password = document.getElementById('password').value.trim();
  loginError.textContent = '';
  try {
    const response = await fetch('/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, password })
    });
    const data = await response.json();
    if (!response.ok) {
      loginError.textContent = data.message || 'Невалиден потребител или парола!';
      return;
    }
    sessionStorage.setItem('authToken', data.token);
    loginContainer.style.display = 'none';
    modal.style.display = 'block';
    loadInitialButtons();
  } catch (e) {
    loginError.textContent = 'Възникна грешка. Опитайте отново.';
  }
});

closeModal.addEventListener('click', () => {
  modal.style.display = 'none';
});

window.addEventListener('click', (event) => {
  if (event.target === modal) {
    modal.style.display = 'none';
  }
});

let buttonsInitialized = false;
function loadInitialButtons() {
  if (buttonsInitialized) return;
  buttonsInitialized = true;
  addAdminButton('Бутон 1');
  addAdminButton('Бутон 2');
  addAdminButton('Бутон 3');
}

function addAdminButton(text) {
  const box = document.createElement('div');
  box.className = 'button-box';

  const btn = document.createElement('button');
  btn.textContent = text;
  box.appendChild(btn);

  const label = document.createElement('label');
  label.style.marginLeft = '10px';
  const checkbox = document.createElement('input');
  checkbox.type = 'checkbox';
  checkbox.checked = true;
  checkbox.addEventListener('change', () => {
    btn.classList.toggle('hidden', !checkbox.checked);
  });
  label.appendChild(checkbox);
  label.appendChild(document.createTextNode(' Показвай'));

  const removeBtn = document.createElement('button');
  removeBtn.textContent = 'Изтрий';
  removeBtn.style.marginLeft = '10px';
  removeBtn.addEventListener('click', () => {
    buttonsArea.removeChild(box);
  });

  box.appendChild(label);
  box.appendChild(removeBtn);
  buttonsArea.appendChild(box);
}

addBtn.addEventListener('click', () => {
  const text = newBtnText.value.trim();
  if (text) {
    addAdminButton(text);
    newBtnText.value = '';
  } else {
    alert('Въведете текст за бутона!');
  }
});
