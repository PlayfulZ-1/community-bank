// ==================== Google OAuth 2.0 Authentication ====================

const Auth = (() => {
  let tokenClient = null;
  let accessToken = null;
  let tokenExpiry = null;
  let currentUser = null;

  async function init() {
    return new Promise((resolve) => {
      if (typeof google === 'undefined' || !google.accounts) {
        console.error('Google Identity Services not loaded');
        resolve(false);
        return;
      }

      tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: CONFIG.CLIENT_ID,
        scope: CONFIG.SCOPES,
        callback: (response) => {
          if (response.error) {
            console.error('Token error:', response.error);
            showToast('เกิดข้อผิดพลาดในการเข้าสู่ระบบ: ' + response.error, 'error');
            return;
          }
          accessToken = response.access_token;
          tokenExpiry = Date.now() + (response.expires_in * 1000) - 60000;
          gapi.client.setToken({ access_token: accessToken });
          _fetchUserInfo();
        },
        error_callback: (err) => {
          console.error('Auth error:', err);
          if (err.type !== 'popup_closed') {
            showToast('เกิดข้อผิดพลาดในการเข้าสู่ระบบ', 'error');
          }
        }
      });

      resolve(true);
    });
  }

  async function _fetchUserInfo() {
    try {
      const resp = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
        headers: { Authorization: `Bearer ${accessToken}` }
      });
      const info = await resp.json();
      currentUser = {
        name: info.name || info.email,
        email: info.email,
        picture: info.picture
      };

      const storedRole = localStorage.getItem('cb_role') || 'committee';
      _saveSession(storedRole);
      _renderUserUI();
      window.AppMain && AppMain.init();
    } catch (e) {
      console.error('Failed to fetch user info:', e);
      showToast('ไม่สามารถดึงข้อมูลผู้ใช้ได้', 'error');
    }
  }

  function _saveSession(role) {
    localStorage.setItem('cb_user', JSON.stringify(currentUser));
    localStorage.setItem('cb_role', role);
  }

  function _renderUserUI() {
    const nameEl = document.getElementById('userDisplayName');
    const roleEl = document.getElementById('userRole');
    const avatarEl = document.getElementById('userAvatar');
    const loginScreen = document.getElementById('loginScreen');
    const mainApp = document.getElementById('mainApp');

    if (nameEl) nameEl.textContent = currentUser.name;
    if (roleEl) {
      const role = getRole();
      roleEl.textContent = role === 'admin' ? 'ผู้ดูแลระบบ' : 'กรรมการ';
    }
    if (avatarEl && currentUser.picture) {
      avatarEl.src = currentUser.picture;
      avatarEl.style.display = 'block';
    }
    if (loginScreen) loginScreen.classList.add('hidden');
    if (mainApp) mainApp.classList.remove('hidden');
  }

  function login() {
    if (!tokenClient) {
      showToast('ระบบยังไม่พร้อม กรุณารอสักครู่', 'warning');
      return;
    }
    tokenClient.requestAccessToken({ prompt: 'select_account' });
  }

  function logout() {
    if (accessToken) {
      google.accounts.oauth2.revoke(accessToken, () => {});
    }
    accessToken = null;
    tokenExpiry = null;
    currentUser = null;
    localStorage.removeItem('cb_user');
    localStorage.removeItem('cb_role');
    gapi.client.setToken(null);
    Cache.clear();

    const mainApp = document.getElementById('mainApp');
    const loginScreen = document.getElementById('loginScreen');
    if (mainApp) mainApp.classList.add('hidden');
    if (loginScreen) loginScreen.classList.remove('hidden');
    showToast('ออกจากระบบแล้ว', 'info');
  }

  async function ensureToken() {
    if (accessToken && tokenExpiry && Date.now() < tokenExpiry) return true;
    if (tokenClient) {
      return new Promise((resolve) => {
        const orig = tokenClient.callback;
        tokenClient.callback = (resp) => {
          orig && orig(resp);
          resolve(!resp.error);
        };
        tokenClient.requestAccessToken({ prompt: '' });
      });
    }
    return false;
  }

  function isAuthenticated() {
    return !!accessToken && Date.now() < (tokenExpiry || 0);
  }

  function getUser() {
    return currentUser;
  }

  function getRole() {
    return localStorage.getItem('cb_role') || 'committee';
  }

  function isAdmin() {
    return getRole() === 'admin';
  }

  function setRole(role) {
    localStorage.setItem('cb_role', role);
    const roleEl = document.getElementById('userRole');
    if (roleEl) roleEl.textContent = role === 'admin' ? 'ผู้ดูแลระบบ' : 'กรรมการ';
    window.AppMain && AppMain.applyRolePermissions();
  }

  function getToken() {
    return accessToken;
  }

  // Attempt to restore session from previous token (silent)
  async function tryRestoreSession() {
    const storedUser = localStorage.getItem('cb_user');
    if (!storedUser) return false;

    try {
      currentUser = JSON.parse(storedUser);
      // Try silent token refresh
      if (tokenClient) {
        return new Promise((resolve) => {
          const timeout = setTimeout(() => resolve(false), 5000);
          const orig = tokenClient.callback;
          tokenClient.callback = (resp) => {
            clearTimeout(timeout);
            orig && orig(resp);
            if (!resp.error) {
              _renderUserUI();
              resolve(true);
            } else {
              resolve(false);
            }
          };
          tokenClient.requestAccessToken({ prompt: '' });
        });
      }
    } catch (e) {
      console.warn('Session restore failed:', e);
    }
    return false;
  }

  return {
    init,
    login,
    logout,
    ensureToken,
    isAuthenticated,
    getUser,
    getRole,
    isAdmin,
    setRole,
    getToken,
    tryRestoreSession
  };
})();
