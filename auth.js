// Inicializar Firebase
firebase.initializeApp(firebaseConfig);
const db = firebase.firestore();

// Elementos DOM
const loginForm = document.getElementById('login-form');
const loginInput = document.getElementById('login');
const loginError = document.getElementById('login-error');

// Manejar login
loginForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    
    const login = loginInput.value.trim().toUpperCase();
    if (!/^[A-Z0-9]{5}$/.test(login)) {
        showError('Login debe ser 5 caracteres alfanuméricos mayúsculas');
        return;
    }

    try {
        const doc = await db.collection('empleados').doc(login).get();
        if (!doc.exists) {
            throw new Error('Usuario no encontrado');
        }

        const userData = { login, ...doc.data() };
        
        // Guardar en sessionStorage
        sessionStorage.setItem('vacaciones_user', JSON.stringify(userData));
        
        // Redirigir según tipo de usuario
        if (userData.admin === 1 || userData.admin === '1') {
            window.location.href = 'admin.html';
        } else {
            window.location.href = 'calendar.html';
        }
    } catch (err) {
        showError(err.message || 'Error de conexión');
    }
});

function showError(message) {
    loginError.textContent = message;
    loginError.style.display = 'block';
}

// Verificar sesión existente al cargar la página
window.addEventListener('DOMContentLoaded', () => {
    const userData = sessionStorage.getItem('vacaciones_user');
    if (userData) {
        const user = JSON.parse(userData);
        if (user.admin === 1 || user.admin === '1') {
            window.location.href = 'admin.html';
        } else {
            window.location.href = 'calendar.html';
        }
    }
});