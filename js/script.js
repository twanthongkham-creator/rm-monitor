// ---- CONFIG ----
const SUPABASE_URL = "https://bdjyxkkzbbzlmxszmvhx.supabase.co";
const SUPABASE_KEY = "sb_publishable_inYG_le-QyiIvjkaUHXyfQ_Nvm4FpR2";
const sb = supabase.createClient(SUPABASE_URL, SUPABASE_KEY);

// ---- POWER AUTOMATE CONFIG (ส่ง Email ผ่าน Outlook 365) ----
// วิธีตั้งค่า (ทำครั้งเดียว ~10 นาที — ไม่ต้องขอ IT Admin):
//   1. ไปที่ https://make.powerautomate.com → สร้าง Flow ใหม่
//   2. เลือก "Instant cloud flow" → Trigger: "When an HTTP request is received" → Create
//   3. ใส่ Request Body JSON Schema ด้านล่างนี้ในช่อง schema ของ Trigger:
//      {"type":"object","properties":{"to":{"type":"string"},"cc":{"type":"string"},"subject":{"type":"string"},"body":{"type":"string"}}}
//   4. กด "+ New step" → ค้นหา "Send an email (V2)" → เลือก Outlook 365
//      - To    : ใส่ Dynamic content → เลือก "to"
//      - Subject: ใส่ Dynamic content → เลือก "subject"
//      - Body  : ใส่ Dynamic content → เลือก "body" → เปิด Code View แล้วเช็คว่าเป็น HTML
//      - CC    : (ขยาย Advanced) → ใส่ Dynamic content → เลือก "cc"
//   5. กด "+ New step" → ค้นหา "Response" → ใส่:
//      - Status Code : 200
//      - Headers     : Content-Type = application/json
//                      Access-Control-Allow-Origin = *
//      - Body        : {"status":"sent"}
//   6. Save Flow → คัดลอก "HTTP POST URL" จาก Trigger มาวางด้านล่าง
const POWER_AUTOMATE_URL = 'https://defaulted45b7d8de16404583789f05db3573.fc.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/c847f4864a5b41b5af6eaad7921b6b2b/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=f-XlkQXp7YKDgDNF_f4YWgQky8vDqbnnM5mjAZUGkG8';

async function sendEmailViaPowerAutomate() {
  const to      = (document.getElementById('email-to')?.value      || '').trim();
  const cc      = (document.getElementById('email-cc')?.value      || '').trim();
  const subject = (document.getElementById('email-subject')?.value || '').trim();
  const bodyEl  = document.getElementById('email-body-preview');

  if (!to)      { showToast('กรุณาระบุอีเมลผู้รับ (To)', 'error');  return; }
  if (!subject) { showToast('กรุณาระบุหัวข้ออีเมล', 'error');        return; }

  if (!POWER_AUTOMATE_URL || POWER_AUTOMATE_URL === 'YOUR_FLOW_URL_HERE') {
    showToast('ยังไม่ได้ตั้งค่า Power Automate URL ใน script.js', 'error');
    return;
  }

  const sendBtn = document.getElementById('graph-send-btn');
  if (sendBtn) {
    sendBtn.disabled = true;
    sendBtn.dataset.originalHtml = sendBtn.innerHTML;
    sendBtn.innerHTML = `<div class="spinner" style="width:12px;height:12px;border-width:1.5px;border-top-color:#fff;border-right-color:transparent;border-bottom-color:transparent;border-left-color:transparent;margin-right:6px;display:inline-block;vertical-align:middle;"></div> กำลังส่ง...`;
  }

  try {
    const fullBody = `<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
      <body style="font-family:'Segoe UI',-apple-system,sans-serif;color:#334155;margin:0;padding:16px 20px;">
        ${bodyEl ? bodyEl.innerHTML : ''}
      </body></html>`;

    await fetch(POWER_AUTOMATE_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ to, cc, subject, body: fullBody })
    });

    showToast('ส่งอีเมล์เรียบร้อยแล้ว! ✉️', 'success');
    closeEmailPreviewModal();
  } catch (err) {
    console.error('Power Automate send error:', err);
    showToast('ส่งอีเมล์ไม่สำเร็จ: ' + (err.message || err), 'error');
  } finally {
    if (sendBtn) {
      sendBtn.disabled = false;
      sendBtn.innerHTML = sendBtn.dataset.originalHtml;
    }
  }
}

// ---- STATE ----
const state = {
  allStock: [],
  allMaster: [],
  allHistory: [],
  stockFiltered: [],
  expiryFiltered: [],
  masterFiltered: [],
  historyFiltered: [],
  agingData: [],
  groups: [],
  stockPage: 1,
  expiryPage: 1,
  masterPage: 1,
  historyPage: 1,
  perPage: 50,
  editId: null,
  stockSortField: 'days_remaining',
  stockSortAsc: true,
  masterSortField: 'code',
  masterSortAsc: true,
  historySortField: 'check_date',
  historySortAsc: false,
  masterColWidths: {},
  stockColWidths: {},
  historyColWidths: {},
  deductColWidths: {},
};

// ---- AUTH FUNCTIONS ----
// ---- AUTH FUNCTIONS ----
async function showLoginOverlay() {
  const overlay = document.createElement('div');
  overlay.className = 'auth-overlay';
  overlay.id = 'global-auth-overlay';
  
  let usersHtml = '';
  let offlineFallback = false;
  
  try {
    const { data, error } = await sb.from('users').select('*').order('name');
    if (error) throw error;
    
    const users = data || [];
    window.loginUsers = users;
    
    if (users.length === 0) {
      offlineFallback = true;
    } else {
      usersHtml = users.map(u => {
        const roleClass = u.role === 'admin' ? 'admin' : (u.role === 'operator' ? 'operator' : 'viewer');
        const roleLabel = u.role === 'admin' ? 'Admin (ผู้ดูแลระบบ)' : (u.role === 'operator' ? 'Operator (ผู้บันทึก)' : 'Viewer (ผู้ดูข้อมูล)');
        const avatarStyle = u.role === 'admin' 
          ? 'background: rgba(79, 70, 229, 0.08); color: #4f46e5; border-color: rgba(79, 70, 229, 0.15);' 
          : (u.role === 'operator' 
            ? 'background: rgba(16, 185, 129, 0.08); color: #10b981; border-color: rgba(16, 185, 129, 0.15);' 
            : 'background: rgba(100, 116, 139, 0.08); color: #64748b; border-color: rgba(100, 116, 139, 0.15);');
        return `
          <div class="auth-item" onclick="handleUserSelect('${u.email}')">
            <div class="auth-avatar" style="${avatarStyle}">
              <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2.5" style="width:15px; height:15px;">
                <path stroke-linecap="round" stroke-linejoin="round" d="M15.75 6a3.75 3.75 0 11-7.5 0 3.75 3.75 0 017.5 0zM4.501 20.118a7.5 7.5 0 0114.998 0A17.933 17.933 0 0112 21.75c-2.676 0-5.216-.584-7.499-1.632z" />
              </svg>
            </div>
            <div class="auth-info">
              <span class="auth-name">${u.name || '—'}</span>
              <span class="auth-email">${u.email}</span>
            </div>
            <span class="auth-badge ${roleClass}">${roleLabel}</span>
          </div>
        `;
      }).join('');
    }
  } catch (err) {
    console.error("Auth error:", err);
    offlineFallback = true;
  }
  
  if (offlineFallback) {
    // Setup offline fallback mock accounts
    window.loginUsers = [
      { email: 'admin@offline.com', name: 'Admin (Offline)', role: 'admin', password: 'admin' },
      { email: 'operator@offline.com', name: 'Operator (Offline)', role: 'operator', password: 'operator' },
      { email: 'viewer@offline.com', name: 'Viewer (Offline)', role: 'viewer', password: '' }
    ];
    
    usersHtml = `
      <div style="padding: 10px; background: rgba(220,38,38,0.05); border: 1px solid rgba(220,38,38,0.15); border-radius: 8px; color: var(--red); font-size:12px; margin-bottom:12px; text-align:center; width:100%;">
        ⚠️ ไม่สามารถเชื่อมต่อฐานข้อมูลผู้ใช้งานได้ (โหมดออฟไลน์ - รหัสผ่านสำหรับ Admin คือ 'admin' และ Operator คือ 'operator')
      </div>
      <div class="auth-item" onclick="handleUserSelect('admin@offline.com')">
        <div class="auth-avatar" style="background: rgba(79, 70, 229, 0.08); color: #4f46e5; border-color: rgba(79, 70, 229, 0.15);">
          <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2.5" style="width:15px; height:15px;">
            <path stroke-linecap="round" stroke-linejoin="round" d="M15.75 6a3.75 3.75 0 11-7.5 0 3.75 3.75 0 017.5 0zM4.501 20.118a7.5 7.5 0 0114.998 0A17.933 17.933 0 0112 21.75c-2.676 0-5.216-.584-7.499-1.632z" />
          </svg>
        </div>
        <div class="auth-info">
          <span class="auth-name">Admin (สิทธิ์ดูแลระบบ)</span>
          <span class="auth-email">เข้าทดสอบระบบในโหมดออฟไลน์</span>
        </div>
        <span class="auth-badge admin">Admin</span>
      </div>
      <div class="auth-item" onclick="handleUserSelect('operator@offline.com')">
        <div class="auth-avatar" style="background: rgba(16, 185, 129, 0.08); color: #10b981; border-color: rgba(16, 185, 129, 0.15);">
          <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2.5" style="width:15px; height:15px;">
            <path stroke-linecap="round" stroke-linejoin="round" d="M15.75 6a3.75 3.75 0 11-7.5 0 3.75 3.75 0 017.5 0zM4.501 20.118a7.5 7.5 0 0114.998 0A17.933 17.933 0 0112 21.75c-2.676 0-5.216-.584-7.499-1.632z" />
          </svg>
        </div>
        <div class="auth-info">
          <span class="auth-name">Operator (สิทธิ์บันทึกข้อมูล)</span>
          <span class="auth-email">เข้าทดสอบระบบในโหมดออฟไลน์</span>
        </div>
        <span class="auth-badge operator">Operator</span>
      </div>
      <div class="auth-item" onclick="handleUserSelect('viewer@offline.com')">
        <div class="auth-avatar" style="background: rgba(100, 116, 139, 0.08); color: #64748b; border-color: rgba(100, 116, 139, 0.15);">
          <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2.5" style="width:15px; height:15px;">
            <path stroke-linecap="round" stroke-linejoin="round" d="M15.75 6a3.75 3.75 0 11-7.5 0 3.75 3.75 0 017.5 0zM4.501 20.118a7.5 7.5 0 0114.998 0A17.933 17.933 0 0112 21.75c-2.676 0-5.216-.584-7.499-1.632z" />
          </svg>
        </div>
        <div class="auth-info">
          <span class="auth-name">Viewer (สิทธิ์ดูข้อมูลทั่วไป)</span>
          <span class="auth-email">เข้าทดสอบระบบในโหมดออฟไลน์</span>
        </div>
        <span class="auth-badge viewer">Viewer</span>
      </div>
    `;
  }
  
  overlay.innerHTML = `
    <div class="auth-card">
      <div class="auth-logo">
        <svg class="logo-svg" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2.5" style="width:28px; height:28px;">
          <path stroke-linecap="round" stroke-linejoin="round" d="M19.428 15.428a2 2 0 00-1.022-.547l-2.387-.477a6 6 0 00-3.86.517l-.318.158a6 6 0 01-3.86.517L6.05 15.39a2 2 0 00-1.806.547M8 4h8l-1 1v5.172a2 2 0 00.586 1.414l5 5c1.26 1.26.367 3.414-1.415 3.414H4.828c-1.782 0-2.674-2.154-1.414-3.414l5-5A2 2 0 009 10.172V5L8 4z"/>
        </svg>
        <h2>StockFlow</h2>
      </div>
      <div class="auth-subtitle">ระบบติดตามสต็อกและวันหมดอายุหัวเชื้อ</div>
      
      <div style="font-weight:600; color:var(--text); align-self:flex-start; margin-bottom:10px; font-size:13px;">เลือกผู้ใช้งานเพื่อเข้าสู่ระบบ:</div>
      <div class="auth-list">
        ${usersHtml}
      </div>
      
      <div class="auth-divider">
        <span>หรือเข้าใช้งานทั่วไป</span>
      </div>
      
      <div class="auth-item" onclick="loginAs('guest@company.com', 'Viewer ทั่วไป', 'viewer')" style="margin-top:8px; border-style:dashed;">
        <div class="auth-avatar" style="background: rgba(100, 116, 139, 0.08); color: #64748b; border-color: rgba(100, 116, 139, 0.15);">
          <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2.5" style="width:15px; height:15px;">
            <path stroke-linecap="round" stroke-linejoin="round" d="M15.75 6a3.75 3.75 0 11-7.5 0 3.75 3.75 0 017.5 0zM4.501 20.118a7.5 7.5 0 0114.998 0A17.933 17.933 0 0112 21.75c-2.676 0-5.216-.584-7.499-1.632z" />
          </svg>
        </div>
        <div class="auth-info">
          <span class="auth-name">Viewer ทั่วไป (Guest)</span>
          <span class="auth-email">ดูข้อมูลทั่วไป ไม่สามารถแก้ไขข้อมูลได้</span>
        </div>
        <span class="auth-badge viewer">Viewer</span>
      </div>
    </div>
  `;
  
  document.body.appendChild(overlay);
}

function handleUserSelect(email) {
  const u = (window.loginUsers || []).find(user => user.email === email);
  if (!u) return;
  
  if (u.role === 'admin' || u.role === 'operator') {
    showPasswordForm(u);
  } else {
    loginAs(u.email, u.name, u.role);
  }
}

function showPasswordForm(u) {
  const card = document.querySelector('.auth-card');
  if (!card) return;
  
  if (!card.dataset.originalHtml) {
    card.dataset.originalHtml = card.innerHTML;
  }
  
  const roleClass = u.role === 'admin' ? 'admin' : (u.role === 'operator' ? 'operator' : 'viewer');
  const roleLabel = u.role === 'admin' ? 'Admin (ผู้ดูแลระบบ)' : (u.role === 'operator' ? 'Operator (ผู้บันทึก)' : 'Viewer (ผู้ดูข้อมูล)');
  const avatarStyle = u.role === 'admin' 
    ? 'background: rgba(79, 70, 229, 0.08); color: #4f46e5; border-color: rgba(79, 70, 229, 0.15);' 
    : 'background: rgba(16, 185, 129, 0.08); color: #10b981; border-color: rgba(16, 185, 129, 0.15);';
  
  card.innerHTML = `
    <div style="align-self: flex-start; margin-bottom: 16px;">
      <button class="btn btn-ghost btn-sm" onclick="cancelPasswordForm()" style="padding: 4px 8px; font-size:12px; display:flex; align-items:center; gap:4px; height:auto; border: 1px solid var(--border);">
        <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" width="12" height="12" stroke-width="2.5">
          <path stroke-linecap="round" stroke-linejoin="round" d="M10.5 19.5L3 12m0 0l7.5-7.5M3 12h18" />
        </svg>
        ย้อนกลับ
      </button>
    </div>
    
    <div style="display: flex; flex-direction: column; align-items: center; gap: 8px; margin-bottom: 20px; width:100%;">
      <div style="width: 54px; height: 54px; border-radius: 50%; display: flex; align-items: center; justify-content: center; ${avatarStyle}">
        <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2.5" style="width: 22px; height: 22px;">
          <path stroke-linecap="round" stroke-linejoin="round" d="M15.75 6a3.75 3.75 0 11-7.5 0 3.75 3.75 0 017.5 0zM4.501 20.118a7.5 7.5 0 0114.998 0A17.933 17.933 0 0112 21.75c-2.676 0-5.216-.584-7.499-1.632z" />
        </svg>
      </div>
      <h3 style="font-size: 15px; font-weight: 600; margin: 0; color: var(--text);">${u.name || '—'}</h3>
      <span style="font-size: 11.5px; color: var(--muted);">${u.email}</span>
      <span class="auth-badge ${roleClass}" style="margin-top: 4px;">${roleLabel}</span>
    </div>
    
    <div style="width: 100%; display: flex; flex-direction: column; gap: 8px;">
      <label style="font-weight:600; font-size:12.5px; color:var(--text);">ป้อนรหัสผ่านเพื่อเข้าใช้งาน <span style="color:var(--red)">*</span></label>
      <input type="password" id="login-password-input" placeholder="ใส่รหัสผ่าน..." style="width: 100%; height: 38px; border: 1px solid var(--border); border-radius: 8px; padding: 8px 12px; font-size: 14px; outline:none; transition:border-color 0.2s;" autofocus>
    </div>
    
    <button class="btn btn-primary" onclick="submitPasswordLogin('${u.email}')" style="width: 100%; height: 38px; justify-content: center; margin-top: 20px;">
      <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" width="14" height="14" stroke-width="2.5" style="margin-right: 6px;">
        <path stroke-linecap="round" stroke-linejoin="round" d="M16.5 10.5V6.75a4.5 4.5 0 10-9 0v3.75m-.75 11.25h10.5a2.25 2.25 0 002.25-2.25v-6.75a2.25 2.25 0 00-2.25-2.25H6.75a2.25 2.25 0 00-2.25 2.25v6.75a2.25 2.25 0 002.25 2.25z" />
      </svg>
      ยืนยันเข้าสู่ระบบ
    </button>
  `;
  
  setTimeout(() => {
    const input = document.getElementById('login-password-input');
    if (input) {
      input.focus();
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') submitPasswordLogin(u.email);
      });
    }
  }, 100);
}

function cancelPasswordForm() {
  const card = document.querySelector('.auth-card');
  if (card && card.dataset.originalHtml) {
    card.innerHTML = card.dataset.originalHtml;
    delete card.dataset.originalHtml;
  }
}

function submitPasswordLogin(email) {
  const u = (window.loginUsers || []).find(user => user.email === email);
  if (!u) return;
  
  const entered = document.getElementById('login-password-input')?.value || '';
  const expected = u.password || '';
  
  if (entered === expected) {
    loginAs(u.email, u.name, u.role);
  } else {
    showToast('❌ รหัสผ่านไม่ถูกต้อง กรุณาลองอีกครั้ง', 'error');
  }
}

function loginAs(email, name, role) {
  localStorage.setItem('currentUserEmail', email);
  localStorage.setItem('currentUserName', name);
  localStorage.setItem('currentUserRole', role);
  location.reload();
}

function logoutUser() {
  localStorage.removeItem('currentUserEmail');
  localStorage.removeItem('currentUserName');
  localStorage.removeItem('currentUserRole');
  location.reload();
}

function insertSidebarUser(name, role) {
  const sidebar = document.querySelector('.sidebar');
  if (!sidebar) return;
  
  // Prevent duplicate cards
  if (document.querySelector('.sidebar-user')) return;
  
  const userPanel = document.createElement('div');
  userPanel.className = 'sidebar-user';
  userPanel.style.padding = '14px 20px';
  userPanel.style.borderTop = '1px solid var(--sidebar-border)';
  userPanel.style.borderBottom = '1px solid var(--sidebar-border)';
  userPanel.style.display = 'flex';
  userPanel.style.flexDirection = 'column';
  userPanel.style.gap = '10px';
  
  const initials = (name || 'U').charAt(0).toUpperCase();
  const roleMap = {
    'admin': 'Admin (ผู้ดูแลระบบ)',
    'operator': 'Operator (ผู้บันทึก)',
    'viewer': 'Viewer (ผู้ดูข้อมูล)'
  };
  const roleLabel = roleMap[role] || role || '—';
  
  userPanel.innerHTML = `
    <div style="display: flex; align-items: center; gap: 10px;">
      <div style="width: 32px; height: 32px; border-radius: 50%; background: var(--sidebar-hover-bg); display: flex; align-items: center; justify-content: center; font-weight: 600; color: var(--sidebar-active-text); font-size: 14px; border:1px solid rgba(255,255,255,0.05); flex-shrink:0;">
        ${initials}
      </div>
      <div style="display: flex; flex-direction: column; overflow: hidden; min-width:0;">
        <span style="font-size: 12px; font-weight: 600; color: #ffffff; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${name}">${name}</span>
        <span style="font-size: 10.5px; color: var(--sidebar-active-text); font-weight:500;" id="user-display-role">${roleLabel}</span>
      </div>
    </div>
    <button onclick="logoutUser()" class="btn" style="width: 100%; justify-content: center; color: #fca5a5; font-size: 11px; padding: 4px; height: 26px; border: 1px dashed rgba(239,68,68,0.3); background:transparent; cursor:pointer; font-weight:500; border-radius:6px; display:flex; align-items:center; gap:4px; transition:all 0.2s;">
      <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" width="12" height="12" stroke-width="2.5">
        <path stroke-linecap="round" stroke-linejoin="round" d="M15.75 9V5.25A2.25 2.25 0 0013.5 3h-6a2.25 2.25 0 00-2.25 2.25v13.5A2.25 2.25 0 007.5 21h6a2.25 2.25 0 002.25-2.25V15M12 9l-3 3m0 0l3 3m-3-3h12.75" />
      </svg>
      ออกจากระบบ / สลับบัญชี
    </button>
  `;
  
  const footer = document.querySelector('.sidebar-footer');
  if (footer) {
    sidebar.insertBefore(userPanel, footer);
  } else {
    sidebar.appendChild(userPanel);
  }
}

function applyRoleRestrictions(role) {
  if (role === 'viewer') {
    // 1. Hide Sidebar items
    const forbiddenPages = ['users.html', 'master.html', 'deduct.html', 'entry.html'];
    forbiddenPages.forEach(p => {
      const link = document.querySelector(`.sidebar a[href="${p}"]`);
      if (link) {
        link.style.display = 'none';
      }
    });
    
    // Hide Section headers if all sub-items are hidden
    const sections = document.querySelectorAll('.sidebar .nav-section');
    sections.forEach(sec => {
      if (sec.textContent.trim() === 'ข้อมูล') {
        sec.style.display = 'none';
      }
    });
    
    // Hide upload and Excel actions if on stock page
    const activePage = document.body.dataset.page;
    if (activePage === 'stock') {
      const excelBtn = document.querySelector('button[onclick="triggerExcelUpload()"]');
      if (excelBtn) excelBtn.style.display = 'none';
    }
    
    // 2. Guard access to forbidden pages (redirection)
    const forbiddenPageIds = ['users', 'master', 'entry', 'deduct'];
    if (forbiddenPageIds.includes(activePage)) {
      window.location.href = 'index.html';
    }
  }
}

// ---- INIT ----
document.addEventListener('DOMContentLoaded', async () => {
  // 1. Verify User Session
  const currentUserEmail = localStorage.getItem('currentUserEmail');
  const currentUserRole = localStorage.getItem('currentUserRole');
  const currentUserName = localStorage.getItem('currentUserName');
  
  if (!currentUserEmail || !currentUserRole) {
    await showLoginOverlay();
    return;
  }
  
  // 2. Insert User profile card in Sidebar
  insertSidebarUser(currentUserName, currentUserRole);
  
  // 3. Apply role restrictions (sidebar hiding & page guarding)
  applyRoleRestrictions(currentUserRole);

  // Apply saved sidebar collapsed state on desktop
  if (window.innerWidth > 768 && localStorage.getItem('sidebar-collapsed') === 'true') {
    document.body.classList.add('sidebar-collapsed');
  }

  const pageId = document.body.dataset.page;
  
  // Highlight active nav item
  if (pageId) {
    const navItem = document.querySelector(`.nav-item[data-page="${pageId}"]`);
    if (navItem) {
      document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
      navItem.classList.add('active');
    }
  }

  // Display today's date
  const todayDateEl = document.getElementById('today-date');
  if (todayDateEl) {
    // ใช้ local date เพื่อป้องกัน UTC offset ทำให้วันที่ผิดช่วง midnight-7am (Thai timezone)
    const now = new Date();
    const localDateStr = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}`;
    todayDateEl.textContent = fmtDate(localDateStr);
  }

  // Set default dates on entry form
  const eCheck = document.getElementById('e-check');
  if (eCheck) eCheck.value = new Date().toISOString().split('T')[0];
  const eReceived = document.getElementById('e-received');
  if (eReceived) eReceived.value = new Date().toISOString().split('T')[0];

  if (typeof initFlatpickr === 'function') initFlatpickr();

  try {
    await loadMaster();
    await loadStock();
    const statusEl = document.getElementById('conn-status');
    if (statusEl) statusEl.textContent = '🟢 เชื่อมต่อแล้ว';
    
    // Auto-render current page
    renderCurrentPage(pageId);
  } catch(e) {
    const statusEl = document.getElementById('conn-status');
    if (statusEl) statusEl.textContent = '🔴 เชื่อมต่อล้มเหลว';
    console.error(e);
  }
});

async function refreshAll() {
  await loadMaster();
  await loadStock();
  const pageId = document.body.dataset.page;
  renderCurrentPage(pageId);
}

function renderCurrentPage(pageId) {
  if (pageId === 'dashboard') renderDashboard();
  else if (pageId === 'expiry') renderExpiry();
  else if (pageId === 'stock') renderStock();
  else if (pageId === 'master') renderMaster();
  else if (pageId === 'entry') loadRecent();
  else if (pageId === 'deduct') renderDeduct();
  else if (pageId === 'users') renderUsersPage();
  else if (pageId === 'history') loadHistory();
}



async function loadStock() {
  const { data, error } = await sb
    .from('stock_with_expiry')
    .select('*')
    .order('check_date', { ascending: false })
    .order('code');
  
  if (error) throw error;
  
  const codeToProductNamesMap = {};
  const codeToRmCodeMap = {};
  const codeToRecipeStepMap = {};
  state.allMaster.forEach(m => {
    if (!m.code) return;
    if (!codeToProductNamesMap[m.code]) {
      codeToProductNamesMap[m.code] = new Set();
    }
    if (m.product_name) {
      codeToProductNamesMap[m.code].add(m.product_name);
    }
    
    if (!codeToRmCodeMap[m.code]) {
      codeToRmCodeMap[m.code] = new Set();
    }
    if (m.rm_code) {
      codeToRmCodeMap[m.code].add(m.rm_code);
    }
    
    if (!codeToRecipeStepMap[m.code]) {
      codeToRecipeStepMap[m.code] = new Set();
    }
    if (m.recipe_step) {
      codeToRecipeStepMap[m.code].add(m.recipe_step);
    }
  });

  state.allStock = (data || []).map(enrichStock)
    .map(s => {
      const nameSet = codeToProductNamesMap[s.code];
      if (nameSet && nameSet.size > 0) {
        s.product_name = Array.from(nameSet).join(', ');
      }
      s.product_group = s.product_name || 'ทั่วไป';
      
      const rmCodeSet = codeToRmCodeMap[s.code];
      s.rm_code = rmCodeSet && rmCodeSet.size > 0 ? Array.from(rmCodeSet).join('<br>') : '—';
      
      const stepSet = codeToRecipeStepMap[s.code];
      s.recipe_step = stepSet && stepSet.size > 0 ? Array.from(stepSet).join('<br>') : '—';
      
      return s;
    })
    .filter(s => !!s.product_name);
  
  // Populate date filter
  const dates = [...new Set(state.allStock.map(s => s.check_date).filter(Boolean))].sort().reverse();
  ['r-date'].forEach(sid => {
    const sel = document.getElementById(sid);
    if (!sel) return;
    const first = sel.options[0];
    sel.innerHTML = '';
    sel.appendChild(first);
    dates.forEach(d => {
      const opt = document.createElement('option');
      opt.value = d;
      opt.textContent = fmtDate(d);
      sel.appendChild(opt);
    });
  });
}

async function loadMaster() {
  const { data, error } = await sb.from('raw_materials').select('*').order('code');
  if (error) throw error;
  state.allMaster = data || [];
  
  // Populate product name filter
  const seenUnified = new Set();
  const pnames = [];
  state.allMaster.forEach(m => {
    if (!m.product_name) return;
    const unified = getUnifiedProductName(m.product_name);
    if (!seenUnified.has(unified)) {
      seenUnified.add(unified);
      pnames.push(unified);
    }
  });
  pnames.sort((a, b) => a.localeCompare(b, 'th', { sensitivity: 'base' }));

  const sel = document.getElementById('m-product');
  if (sel) {
    const first = sel.options[0];
    sel.innerHTML = '';
    sel.appendChild(first);
    pnames.forEach(p => {
      const opt = document.createElement('option');
      opt.value = p;
      opt.textContent = p;
      sel.appendChild(opt);
    });
  }

  // Populate raw-materials-list datalist in entry.html with unique raw materials
  const eDatalist = document.getElementById('raw-materials-list');
  if (eDatalist && document.body.dataset.page === 'entry') {
    eDatalist.innerHTML = '';
    const seenCodes = new Set();
    state.allMaster.forEach(m => {
      if (!m.code || seenCodes.has(m.code)) return;
      seenCodes.add(m.code);
      const opt = document.createElement('option');
      opt.value = m.code;
      opt.textContent = `${m.code} — ${m.name || m.recipe_step || m.rm_name || ''}`;
      opt.dataset.desc = m.name || '';
      opt.dataset.group = m.product_name || '';
      eDatalist.appendChild(opt);
    });
  }

  // Populate exp-group and st-group filters
  ['exp-group','st-group'].forEach(sid => {
    const sel = document.getElementById(sid);
    if (!sel) return;
    const first = sel.options[0];
    sel.innerHTML = '';
    sel.appendChild(first);
    pnames.forEach(p => {
      const opt = document.createElement('option');
      opt.value = p;
      opt.textContent = p;
      sel.appendChild(opt);
    });
  });
}


function enrichStock(row) {
  const eff = row.extended_expiry_2 || row.extended_expiry_1 || row.expiry_date;
  // เปรียบเทียบ UTC midnight ของวันหมดอายุ กับ UTC midnight ของวันนี้ เพื่อให้ตรงกับ fmtDate
  const todayUTC = Date.UTC(new Date().getUTCFullYear(), new Date().getUTCMonth(), new Date().getUTCDate());
  const days = eff ? Math.round((new Date(eff).getTime() - todayUTC) / 86400000) : null;
  let status = 'ok';
  if (days === null) status = 'unknown';
  else if (days < 0) status = 'expired';
  else if (days <= 30) status = 'critical';
  else if (days <= 60) status = 'warning';
  
  // Look up master name
  let rm_full_name = null;
  if (state.allMaster.length > 0) {
    const match = state.allMaster.find(m => m.code === row.code);
    if (match) rm_full_name = match.name;
  }
  if (!rm_full_name) {
    rm_full_name = row.rm_full_name || row.new_description || row.description;
  }
  
  return { ...row, effective_expiry: eff, days_remaining: days, expiry_status: status, rm_full_name };
}

// ---- DASHBOARD ----
function renderDashboard() {
  const latest = getLatestStock();
  const expired = latest.filter(s => s.expiry_status === 'expired');
  const critical = latest.filter(s => s.expiry_status === 'critical');
  const warning = latest.filter(s => s.expiry_status === 'warning');
  const ok = latest.filter(s => s.expiry_status === 'ok');

  const kpiTotalEl = document.getElementById('kpi-total');
  if (kpiTotalEl) kpiTotalEl.textContent = latest.length;
  const kpiExpiredEl = document.getElementById('kpi-expired');
  if (kpiExpiredEl) kpiExpiredEl.textContent = expired.length;
  const kpiCriticalEl = document.getElementById('kpi-critical');
  if (kpiCriticalEl) kpiCriticalEl.textContent = critical.length;
  const kpiWarningEl = document.getElementById('kpi-warning');
  if (kpiWarningEl) kpiWarningEl.textContent = warning.length;
  const kpiOkEl = document.getElementById('kpi-ok');
  if (kpiOkEl) kpiOkEl.textContent = ok.length;

  // Alerts
  const alerts = [...expired, ...critical];
  const alertsEl = document.getElementById('alerts-container');
  if (alertsEl) {
    if (alerts.length > 0) {
      const itemsHtml = alerts.slice(0, 5).map(a => {
        const nameText = a.rm_full_name || a.new_description || a.description || '—';
        const isExpired = a.days_remaining < 0;
        const statusClass = isExpired ? 'status-expired' : 'status-critical';
        const badgeClass = isExpired ? 'expired' : 'critical';
        const badgeLabel = isExpired ? 'Expired' : 'Critical';
        const timeLabel = isExpired 
          ? `หมดอายุไปแล้ว ${Math.abs(a.days_remaining)} วัน` 
          : `เหลืออีก ${a.days_remaining} วัน`;
        
        return `
          <div class="alert-box-item ${statusClass}">
            <span class="alert-badge ${badgeClass}">${badgeLabel}</span>
            <span class="alert-info" title="${a.code} — ${nameText} [${a.product_name || ''}]">
              <strong>${a.code}</strong> — ${nameText} 
              <span class="alert-batch">Batch: ${a.batch || 'no-batch'}</span>
            </span>
            <span class="alert-time">${timeLabel}</span>
          </div>
        `;
      }).join('');

      alertsEl.innerHTML = `
        <div class="alert-box-card">
          <div class="alert-box-header">
            <span class="alert-box-title">🚨 การแจ้งเตือนวัตถุดิบวิกฤต (Critical Stock Alerts)</span>
            <span class="alert-box-count">พบทั้งหมด ${alerts.length} รายการ (แสดง 5 อันดับแรก)</span>
          </div>
          <div class="alert-box-list">
            ${itemsHtml}
          </div>
        </div>
      `;
    } else {
      alertsEl.innerHTML = '';
    }
  }

  // Table
  const urgent = latest.filter(s => ['expired','critical','warning'].includes(s.expiry_status))
    .sort((a,b) => (a.days_remaining??9999) - (b.days_remaining??9999));
  
  const tbody = document.getElementById('dash-tbody');
  if (tbody) {
    tbody.innerHTML = urgent.length
      ? urgent.slice(0,20).map(r => `<tr class="row-${r.expiry_status}">
          <td><span class="status-dot dot-${r.expiry_status}"></span></td>
          <td class="code-text">${r.code}</td>
          <td class="wrap">${r.rm_full_name || r.new_description || r.description || '—'}</td>
          <td>${(r.product_name || '—').split(', ').join('<br>')}</td>
          <td class="code-text">${r.batch || '—'}</td>
          <td><span class="days-chip days-${r.expiry_status}">${r.days_remaining ?? '?'} วัน</span></td>
          <td>${r.quantity ?? '—'}</td>
        </tr>`).join('')
      : '<tr><td colspan="7" class="empty"><div class="empty-icon">✅</div>ไม่มีรายการที่น่าเป็นห่วง</td></tr>';

  }

  // Donut Chart updates
  const total = latest.length;
  const donutTotalEl = document.getElementById('donut-total');
  if (donutTotalEl) donutTotalEl.textContent = total;

  const expPct = total ? (expired.length / total) * 100 : 0;
  const critPct = total ? (critical.length / total) * 100 : 0;
  const warnPct = total ? (warning.length / total) * 100 : 0;
  const okPct = total ? (ok.length / total) * 100 : 0;

  const donutChartEl = document.querySelector('.donut-chart');
  if (donutChartEl) {
    donutChartEl.style.background = `conic-gradient(
      var(--red) 0% ${expPct}%,
      var(--orange) ${expPct}% ${expPct + critPct}%,
      var(--yellow) ${expPct + critPct}% ${expPct + critPct + warnPct}%,
      var(--green) ${expPct + critPct + warnPct}% 100%
    )`;
  }

  const legExpired = document.getElementById('legend-expired');
  if (legExpired) legExpired.textContent = `${expired.length} (${Math.round(expPct)}%)`;
  const legCritical = document.getElementById('legend-critical');
  if (legCritical) legCritical.textContent = `${critical.length} (${Math.round(critPct)}%)`;
  const legWarning = document.getElementById('legend-warning');
  if (legWarning) legWarning.textContent = `${warning.length} (${Math.round(warnPct)}%)`;
  const legOk = document.getElementById('legend-ok');
  if (legOk) legOk.textContent = `${ok.length} (${Math.round(okPct)}%)`;

  // Product Group Risk Analysis updates
  const groupRiskMap = {};
  latest.forEach(item => {
    const groupName = item.product_name || 'ทั่วไป';
    if (!groupRiskMap[groupName]) {
      groupRiskMap[groupName] = { expired: 0, critical: 0, warning: 0, ok: 0, total: 0 };
    }
    if (item.expiry_status === 'expired') groupRiskMap[groupName].expired++;
    else if (item.expiry_status === 'critical') groupRiskMap[groupName].critical++;
    else if (item.expiry_status === 'warning') groupRiskMap[groupName].warning++;
    else groupRiskMap[groupName].ok++;
    groupRiskMap[groupName].total++;
  });

  const groupRiskList = Object.entries(groupRiskMap)
    .map(([name, data]) => {
      const riskCount = data.expired + data.critical + data.warning;
      return { name, riskCount, ...data };
    })
    .filter(g => g.riskCount > 0)
    .sort((a, b) => b.riskCount - a.riskCount);

  const groupRiskListEl = document.getElementById('group-risk-list');
  if (groupRiskListEl) {
    groupRiskListEl.innerHTML = groupRiskList.length
      ? groupRiskList.slice(0, 4).map(group => {
          const pct = (group.riskCount / group.total) * 100;
          return `
            <div class="qty-item">
              <div class="qty-info">
                <div class="qty-name" title="${group.name}">${group.name}</div>
                <div class="qty-code">ความเสี่ยง: ${Math.round(pct)}% ของกลุ่ม</div>
              </div>
              <div class="qty-bar-container" style="width: 130px;">
                <div class="qty-val" style="color: var(--red);">${group.riskCount} Batches เสี่ยง</div>
                <div style="font-size:10px; color:var(--muted); text-align:right;">หมดอายุ: ${group.expired} | วิกฤต: ${group.critical}</div>
              </div>
            </div>
          `;
        }).join('')
      : '<div style="text-align:center;color:var(--muted);padding:20px;">✅ ทุกกลุ่มผลิตภัณฑ์ไม่มีรายการวิกฤต</div>';
  }

  // Stock Aging Distribution updates
  const lt3 = latest.filter(r => {
    const receivedDays = r.received_date ? Math.round((new Date() - new Date(r.received_date)) / 86400000) : null;
    return receivedDays !== null && receivedDays < 90;
  }).length;
  const m36 = latest.filter(r => {
    const receivedDays = r.received_date ? Math.round((new Date() - new Date(r.received_date)) / 86400000) : null;
    return receivedDays !== null && receivedDays >= 90 && receivedDays < 180;
  }).length;
  const gt6 = latest.filter(r => {
    const receivedDays = r.received_date ? Math.round((new Date() - new Date(r.received_date)) / 86400000) : null;
    return receivedDays !== null && receivedDays >= 180;
  }).length;

  const totalAging = lt3 + m36 + gt6 || 1;
  const lt3Pct = (lt3 / totalAging) * 100;
  const m36Pct = (m36 / totalAging) * 100;
  const gt6Pct = (gt6 / totalAging) * 100;

  const ageLt3El = document.getElementById('aging-dist-lt3');
  if (ageLt3El) ageLt3El.textContent = `${lt3} รายการ (${Math.round(lt3Pct)}%)`;
  const ageM36El = document.getElementById('aging-dist-m36');
  if (ageM36El) ageM36El.textContent = `${m36} รายการ (${Math.round(m36Pct)}%)`;
  const ageGt6El = document.getElementById('aging-dist-gt6');
  if (ageGt6El) ageGt6El.textContent = `${gt6} รายการ (${Math.round(gt6Pct)}%)`;

  const barLt3 = document.getElementById('aging-bar-lt3');
  if (barLt3) barLt3.style.width = `${lt3Pct}%`;
  const barM36 = document.getElementById('aging-bar-m36');
  if (barM36) barM36.style.width = `${m36Pct}%`;
  const barGt6 = document.getElementById('aging-bar-gt6');
  if (barGt6) barGt6.style.width = `${gt6Pct}%`;


  // ---- QUANTITY ANALYSIS ----
  // Helper to look up unit from master list
  const getUnit = (code) => {
    const match = state.allMaster.find(m => m.code === code);
    return match ? (match.unit || 'Units') : 'Units';
  };

  // Aggregate by raw material code
  const codeQtyMap = {};
  latest.forEach(item => {
    const code = item.code;
    const qty = parseFloat(item.quantity) || 0;
    const name = item.rm_full_name || item.new_description || item.description || '—';
    if (!codeQtyMap[code]) {
      codeQtyMap[code] = {
        code: code,
        name: name,
        quantity: 0,
        unit: getUnit(code)
      };
    }
    codeQtyMap[code].quantity += qty;
  });

  const aggregatedList = Object.values(codeQtyMap);
  const totalVolume = aggregatedList.reduce((sum, item) => sum + item.quantity, 0);
  const uniqueItemsCount = aggregatedList.length;

  // Render total summary numbers
  const qtyTotalSumEl = document.getElementById('qty-total-sum');
  if (qtyTotalSumEl) {
    qtyTotalSumEl.textContent = totalVolume.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  const qtyTotalItemsEl = document.getElementById('qty-total-items');
  if (qtyTotalItemsEl) {
    qtyTotalItemsEl.textContent = uniqueItemsCount;
  }

  // Sort for Max Stock
  const maxStockList = [...aggregatedList].sort((a, b) => b.quantity - a.quantity).slice(0, 5);
  
  // Sort for Low Stock (excluding 0, since latest only contains quantity > 0)
  const minStockList = [...aggregatedList].sort((a, b) => a.quantity - b.quantity).slice(0, 5);

  // Render Max Stock list
  const maxListEl = document.getElementById('qty-max-list');
  if (maxListEl) {
    const highestQty = maxStockList[0] ? maxStockList[0].quantity : 1;
    maxListEl.innerHTML = maxStockList.length
      ? maxStockList.map(item => {
          const pct = Math.max(5, (item.quantity / highestQty) * 100);
          return `
            <div class="qty-item">
              <div class="qty-info">
                <div class="qty-name" title="${item.name}">${item.name}</div>
                <div class="qty-code">${item.code}</div>
              </div>
              <div class="qty-bar-container">
                <div class="qty-val">${item.quantity.toLocaleString(undefined, { maximumFractionDigits: 2 })} <span style="font-size:11px;color:var(--muted);font-weight:normal;">${item.unit}</span></div>
                <div class="qty-bar">
                  <div class="qty-bar-fill max" style="width: ${pct}%"></div>
                </div>
              </div>
            </div>
          `;
        }).join('')
      : '<div style="text-align:center;color:var(--muted);padding:20px;">ไม่มีข้อมูล</div>';
  }

  // Render Min/Low Stock list
  const minListEl = document.getElementById('qty-min-list');
  if (minListEl) {
    const localHighest = minStockList[minStockList.length - 1] ? minStockList[minStockList.length - 1].quantity : 1;
    minListEl.innerHTML = minStockList.length
      ? minStockList.map(item => {
          const pct = Math.max(5, (item.quantity / localHighest) * 100);
          return `
            <div class="qty-item">
              <div class="qty-info">
                <div class="qty-name" title="${item.name}">${item.name}</div>
                <div class="qty-code">${item.code}</div>
              </div>
              <div class="qty-bar-container">
                <div class="qty-val">${item.quantity.toLocaleString(undefined, { maximumFractionDigits: 2 })} <span style="font-size:11px;color:var(--muted);font-weight:normal;">${item.unit}</span></div>
                <div class="qty-bar">
                  <div class="qty-bar-fill min" style="width: ${pct}%"></div>
                </div>
              </div>
            </div>
          `;
        }).join('')
      : '<div style="text-align:center;color:var(--muted);padding:20px;">ไม่มีข้อมูล</div>';
  }
}
// ---- EXPIRY MONITOR & AGING ----
function filterExpiry() {
  const searchEl = document.getElementById('exp-search');
  const groupEl = document.getElementById('exp-group');
  const statusEl = document.getElementById('exp-status');
  if (!searchEl || !groupEl || !statusEl) return;

  const q = searchEl.value.toLowerCase();
  const grp = groupEl.value;
  const status = statusEl.value;
  
  // Get active unique stock records
  const latest = getLatestStock();
  
  // 1. Calculate overall Aging Cards (all active stock, not just filtered ones)
  const allRowsWithAging = latest.map(r => {
    const receivedDays = r.received_date ? Math.round((new Date() - new Date(r.received_date)) / 86400000) : null;
    let aging = '—';
    if (receivedDays !== null) {
      if (receivedDays < 90) aging = '< 3M';
      else if (receivedDays < 180) aging = '3–6M';
      else aging = '> 6M';
    }
    return { ...r, aging };
  });
  
  const lt3 = allRowsWithAging.filter(r => r.aging === '< 3M').length;
  const m36 = allRowsWithAging.filter(r => r.aging === '3–6M').length;
  const gt6 = allRowsWithAging.filter(r => r.aging === '> 6M').length;
  
  const lt3El = document.getElementById('ag-lt3');
  if (lt3El) lt3El.textContent = lt3;
  const m36El = document.getElementById('ag-36');
  if (m36El) m36El.textContent = m36;
  const gt6El = document.getElementById('ag-gt6');
  if (gt6El) gt6El.textContent = gt6;

  // 2. Filter list according to criteria
  state.expiryFiltered = allRowsWithAging.filter(r => {
    const text = `${r.code} ${r.rm_full_name || ''} ${r.description || ''} ${r.new_description || ''} ${r.batch || ''}`.toLowerCase();
    
    let grpMatch = !grp;
    if (grp && r.product_name) {
      const names = r.product_name.split(',').map(n => getUnifiedProductName(n.trim()));
      grpMatch = names.includes(grp);
    }
    
    return (!q || text.includes(q))
      && grpMatch
      && (!status || r.expiry_status === status);
  }).sort((a,b) => (a.days_remaining??9999) - (b.days_remaining??9999));
  
  state.expiryPage = 1;
  renderExpiryTable();
}

function renderExpiry() {
  filterExpiry();
}

function renderExpiryTable() {
  const data = state.expiryFiltered;
  const start = (state.expiryPage-1) * state.perPage;
  const page = data.slice(start, start+state.perPage);
  
  const countEl = document.getElementById('exp-count');
  if (countEl) countEl.textContent = `${data.length} รายการ`;
  
  const tbody = document.getElementById('exp-tbody');
  if (tbody) {
    tbody.innerHTML = page.length
      ? page.map(r => {
          const unitPerPack = r.unit_per_pack || 0;
          const prodUnits = unitPerPack > 0 ? (r.quantity * unitPerPack).toFixed(0) : '—';
          const nameText = r.rm_full_name || r.new_description || r.description || '—';
          return `<tr class="row-${r.expiry_status}">
            <td style="text-align:center;"><span class="status-dot dot-${r.expiry_status}"></span></td>
            <td class="code-text" title="${r.code}">${r.code}</td>
            <td title="${nameText}">${nameText}</td>
            <td title="${r.product_name || ''}">${(r.product_name || '—').split(', ').join('<br>')}</td>
            <td class="code-text" title="${r.batch || ''}">${r.batch || '—'}</td>
            <td>${fmtDate(r.production_date)}</td>
            <td>${fmtDate(r.expiry_date)}</td>
            <td>
              ${fmtDate(r.extended_expiry_1)}
              ${r.extended_expiry_1_doc ? `<a href="${r.extended_expiry_1_doc}" target="_blank" title="ดูเอกสารต่ออายุครั้งที่ 1" style="margin-left:4px; text-decoration:none; font-size:12px;">📄</a>` : ''}
            </td>
            <td>
              ${fmtDate(r.extended_expiry_2)}
              ${r.extended_expiry_2_doc ? `<a href="${r.extended_expiry_2_doc}" target="_blank" title="ดูเอกสารต่ออายุครั้งที่ 2" style="margin-left:4px; text-decoration:none; font-size:12px;">📄</a>` : ''}
            </td>
            <td style="text-align:center;"><span class="days-chip days-${r.expiry_status}">${r.days_remaining ?? '?'} วัน</span></td>
            <td style="text-align:center;"><span class="badge ${agingBadge(r.aging)}">${r.aging}</span></td>
            <td>${r.quantity ?? '—'}</td>
            <td>${prodUnits}</td>
            <td>${fmtDate(r.check_date)}</td>
            <td>
              ${localStorage.getItem('currentUserRole') === 'viewer' ? '' : `
              <button class="btn btn-ghost btn-sm" onclick="openEdit('${r.id}')"><svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" width="14" height="14" style="stroke-width:2"><path stroke-linecap="round" stroke-linejoin="round" d="M16.862 4.487l1.687-1.688a1.875 1.875 0 112.652 2.652L6.832 19.82a4.5 4.5 0 01-1.897 1.13l-2.685.8.8-2.685a4.5 4.5 0 011.13-1.897L16.863 4.487zm0 0L19.5 7.125" /></svg></button>
              `}
            </td>
          </tr>`;
        }).join('')
      : '<tr><td colspan="15" class="empty"><div class="empty-icon">🔍</div>ไม่พบข้อมูล</td></tr>';
  }
  
  renderPagination('exp', data.length, state.expiryPage, (p) => { state.expiryPage=p; renderExpiryTable(); });
}

// ---- STOCK ALL ----
function filterStock() {
  const searchEl = document.getElementById('st-search');
  const groupEl = document.getElementById('st-group');
  if (!searchEl || !groupEl) return;

  const q = searchEl.value.toLowerCase();
  const grp = groupEl.value;
  state.activeStockGroupFilter = grp;
  
  // 1. Filter raw list (matches can duplicate if an item has multiple formulas)
  const filteredRaw = state.allStock.filter(r => {
    const text = `${r.code} ${r.rm_full_name || ''} ${r.description || ''} ${r.new_description || ''} ${r.batch || ''}`.toLowerCase();
    
    let grpMatch = !grp;
    if (grp && r.product_name) {
      const names = r.product_name.split(',').map(n => getUnifiedProductName(n.trim()));
      grpMatch = names.includes(grp);
    }
    
    return (!q || text.includes(q)) && grpMatch;
  });
  
  // 2. Deduplicate by unique stock entry ID to avoid duplicate batch rows
  const uniqueList = [];
  const seenIds = new Set();
  filteredRaw.forEach(r => {
    if (!seenIds.has(r.id)) {
      seenIds.add(r.id);
      uniqueList.push(r);
    }
  });
  
  state.stockFiltered = uniqueList;
  state.stockPage = 1;
  renderStockTable();
}

function sortStock(field) {
  if (state.stockSortField === field) {
    state.stockSortAsc = !state.stockSortAsc;
  } else {
    state.stockSortField = field;
    state.stockSortAsc = true;
  }
  filterStock();
}

function renderStock() { filterStock(); }

function toggleStockGroup(code) {
  if (!state.expandedStockGroups) {
    state.expandedStockGroups = new Set();
  }
  if (state.expandedStockGroups.has(code)) {
    state.expandedStockGroups.delete(code);
  } else {
    state.expandedStockGroups.add(code);
  }
  renderStockTable();
}

function groupStockData(rawData) {
  const groups = [];
  const groupsMap = new Map();
  
  rawData.forEach(item => {
    const code = item.code;
    if (!groupsMap.has(code)) {
      const codeMatches = state.allMaster.filter(m => m.code === code);
      let targetMatches = codeMatches;
      if (state.activeStockGroupFilter) {
        const filtered = codeMatches.filter(m => getUnifiedProductName(m.product_name) === state.activeStockGroupFilter);
        if (filtered.length > 0) {
          targetMatches = filtered;
        }
      }
      
      const names = [...new Set(targetMatches.map(m => m.product_name).filter(Boolean))];
      const rmCodes = [...new Set(targetMatches.map(m => m.rm_code).filter(Boolean))];
      const steps = [...new Set(targetMatches.map(m => m.recipe_step).filter(Boolean))];

      const g = {
        code: code,
        rm_code: rmCodes.length ? rmCodes.join('<br>') : '—',
        rm_full_name: item.rm_full_name || item.new_description || item.description || '—',
        recipe_step: steps.length ? steps.join('<br>') : '—',
        product_group: names.length ? names.join('<br>') : '—',
        batches: [],
        total_quantity: 0,
        min_days_remaining: Infinity,
        worst_status: 'ok',
        soonest_expiry: null,
        soonest_received: null,
        soonest_production: null,
        representative: item
      };
      groupsMap.set(code, g);
      groups.push(g);
    }
    
    const group = groupsMap.get(code);
    group.batches.push(item);
    
    // Sum quantity
    group.total_quantity += parseFloat(item.quantity) || 0;
    
    // Find min days remaining
    const days = item.days_remaining !== null && item.days_remaining !== undefined ? parseFloat(item.days_remaining) : null;
    if (days !== null) {
      if (days < group.min_days_remaining) {
        group.min_days_remaining = days;
        group.soonest_expiry = item.effective_expiry || item.extended_expiry_2 || item.extended_expiry_1 || item.expiry_date;
        group.soonest_received = item.received_date;
        group.soonest_production = item.production_date;
      }
    }
    
    // Determine worst status: expired > critical > warning > ok > unknown
    const statusPriority = { expired: 4, critical: 3, warning: 2, ok: 1, unknown: 0 };
    const curStatus = item.expiry_status || 'unknown';
    const worstStatus = group.worst_status || 'unknown';
    if (statusPriority[curStatus] > statusPriority[worstStatus]) {
      group.worst_status = curStatus;
    }
  });
  
  // Format min_days_remaining if it is still Infinity
  groups.forEach(group => {
    if (group.min_days_remaining === Infinity) {
      group.min_days_remaining = null;
    }
    // Sort the batches inside each group by expiry so soonest is first
    group.batches.sort((a, b) => {
      const daysA = a.days_remaining !== null ? a.days_remaining : Infinity;
      const daysB = b.days_remaining !== null ? b.days_remaining : Infinity;
      return daysA - daysB;
    });
  });
  
  return groups;
}

function renderStockTable() {
  // Render headers dynamically with sorting icons
  const thead = document.querySelector('#page-stock table thead');
  if (thead) {
    const cols = [
      { label: 'สถานะ', field: 'expiry_status' },
      { label: 'Code', field: 'code' },
      { label: 'RM Code', field: 'rm_code' },
      { label: 'ชื่อ', field: 'rm_full_name' },
      { label: 'ขั้นตอน (Step)', field: 'recipe_step' },
      { label: 'กลุ่ม', field: 'product_group' },
      { label: 'Batch', field: 'batch' },
      { label: 'วันรับ', field: 'received_date' },
      { label: 'วันผลิต', field: 'production_date' },
      { label: 'วันหมดอายุ', field: 'effective_expiry' },
      { label: 'วันที่เหลือ', field: 'days_remaining' },
      { label: 'จำนวน', field: 'quantity' },
      { label: '', field: null }
    ];
    thead.innerHTML = `<tr>${cols.map(c => {
      if (!c.field) return '<th style="text-align:center;"></th>';
      const isSorted = state.stockSortField === c.field;
      const icon = isSorted ? (state.stockSortAsc ? ' ▲' : ' ▼') : ' ↕';
      const savedWidth = state.stockColWidths[c.field];
      const widthStyle = savedWidth ? `width: ${savedWidth}px;` : '';
      return `<th data-field="${c.field}" onclick="sortStock('${c.field}')" class="sortable-th ${isSorted ? 'sorted-active' : ''}" style="cursor:pointer; user-select:none; position:relative; text-align:center; ${widthStyle}">
        <div class="th-content" style="display:inline-flex; align-items:center; gap:4px; justify-content:center; width:100%;">
          ${c.label} <span class="sort-icon" style="font-size:10px; color:${isSorted ? 'var(--primary)' : 'var(--muted)'};">${icon}</span>
        </div>
      </th>`;
    }).join('')}</tr>`;
  }

  // 1. Group the stock data
  const groupedData = groupStockData(state.stockFiltered);
  
  // 2. Sort the grouped data
  if (state.stockSortField) {
    const f = state.stockSortField;
    const asc = state.stockSortAsc ? 1 : -1;
    groupedData.sort((a, b) => {
      let valA, valB;
      if (f === 'expiry_status') {
        valA = a.worst_status;
        valB = b.worst_status;
      } else if (f === 'code') {
        valA = a.code;
        valB = b.code;
      } else if (f === 'rm_code') {
        valA = a.representative.rm_code || '';
        valB = b.representative.rm_code || '';
      } else if (f === 'rm_full_name') {
        valA = a.rm_full_name;
        valB = b.rm_full_name;
      } else if (f === 'recipe_step') {
        valA = a.representative.recipe_step || '';
        valB = b.representative.recipe_step || '';
      } else if (f === 'product_group') {
        valA = a.product_group;
        valB = b.product_group;
      } else if (f === 'batch') {
        valA = a.batches[0] ? a.batches[0].batch : '';
        valB = b.batches[0] ? b.batches[0].batch : '';
      } else if (f === 'received_date') {
        valA = a.soonest_received || '';
        valB = b.soonest_received || '';
      } else if (f === 'production_date') {
        valA = a.soonest_production || '';
        valB = b.soonest_production || '';
      } else if (f === 'effective_expiry') {
        valA = a.soonest_expiry || '';
        valB = b.soonest_expiry || '';
      } else if (f === 'days_remaining') {
        const numA = a.min_days_remaining !== null ? a.min_days_remaining : Infinity;
        const numB = b.min_days_remaining !== null ? b.min_days_remaining : Infinity;
        return (numA - numB) * asc;
      } else if (f === 'quantity') {
        return (a.total_quantity - b.total_quantity) * asc;
      } else {
        valA = a.representative[f] || '';
        valB = b.representative[f] || '';
      }
      
      if (valA === undefined || valA === null) valA = '';
      if (valB === undefined || valB === null) valB = '';
      
      return String(valA).localeCompare(String(valB), 'th', { sensitivity: 'base' }) * asc;
    });
  }

  // 3. Paginate the grouped data
  const start = (state.stockPage-1) * state.perPage;
  const page = groupedData.slice(start, start+state.perPage);
  
  const countEl = document.getElementById('st-count');
  if (countEl) {
    const sumQty = state.stockFiltered.reduce((sum, r) => sum + (parseFloat(r.quantity) || 0), 0);
    countEl.innerHTML = `
      <span style="margin-right: 12px; color: var(--muted); font-size: 13px;">${groupedData.length} วัตถุดิบ (${state.stockFiltered.length} Batches)</span>
      <span style="background-color: #e0f2fe; color: #0369a1; border: 1px solid #bae6fd; padding: 5px 10px; border-radius: 6px; font-weight: 600; font-size: 13px; display: inline-flex; align-items: center; gap: 6px; box-shadow: 0 1px 2px rgba(0,0,0,0.02);">
        <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="2.5" stroke="currentColor" width="13" height="13">
          <path stroke-linecap="round" stroke-linejoin="round" d="M20.25 7.5l-.625 10.632a2.25 2.25 0 01-2.247 2.118H6.622a2.25 2.25 0 01-2.247-2.118L3.75 7.5M10 11.25h4M3.375 7.5h17.25c.621 0 1.125-.504 1.125-1.125v-1.5c0-.621-.504-1.125-1.125-1.125H3.375c-.621 0-1.125.504-1.125 1.125v1.5c0 .621.504 1.125 1.125 1.125z" />
        </svg>
        ยอดรวม: ${sumQty.toLocaleString(undefined, {minimumFractionDigits: 0, maximumFractionDigits: 2})}
      </span>
    `;
  }
  
  if (!state.expandedStockGroups) {
    state.expandedStockGroups = new Set();
  }
  
  const tbody = document.getElementById('st-tbody');
  if (tbody) {
    if (page.length === 0) {
      tbody.innerHTML = '<tr><td colspan="11" class="empty"><div class="empty-icon">📦</div>ไม่พบข้อมูล</td></tr>';
    } else {
      let rowsHtml = '';
      page.forEach(r => {
        const isExpanded = state.expandedStockGroups.has(r.code);
        const icon = isExpanded ? '▼' : '▶';
        
        // Render main row
        rowsHtml += `
          <tr class="row-${r.worst_status} main-group-row" style="font-weight: 500;">
            <td onclick="toggleStockGroup('${r.code}')" style="cursor:pointer; text-align:center;"><span class="status-dot dot-${r.worst_status}"></span></td>
            <td onclick="toggleStockGroup('${r.code}')" class="code-text" style="cursor:pointer; font-weight: 600; color: var(--primary); text-align:center;">
              <span style="display:inline-block; width:12px; margin-right:4px;">${icon}</span>${r.code}
            </td>
            <td onclick="toggleStockGroup('${r.code}')" class="code-text" style="cursor:pointer; text-align:center;">${r.rm_code}</td>
            <td onclick="toggleStockGroup('${r.code}')" class="wrap" style="cursor:pointer; text-align:left;">${r.rm_full_name}</td>
            <td onclick="toggleStockGroup('${r.code}')" style="font-size:12px; cursor:pointer; text-align:center;">${r.recipe_step}</td>
            <td onclick="toggleStockGroup('${r.code}')" style="font-size:12px; cursor:pointer; text-align:center;">${r.product_group.split(', ').join('<br>')}</td>
            <td onclick="toggleStockGroup('${r.code}')" style="cursor:pointer; text-align:center;">
              <span class="badge" style="background-color: var(--surface2); border: 1px solid var(--border); color: var(--text-muted); font-size:11px;">
                ${r.batches.length} Batch${r.batches.length > 1 ? 'es' : ''}
              </span>
            </td>
            <td onclick="toggleStockGroup('${r.code}')" style="cursor:pointer; text-align:center;">${fmtDate(r.soonest_received)}</td>
            <td onclick="toggleStockGroup('${r.code}')" style="cursor:pointer; text-align:center;">${fmtDate(r.soonest_production)}</td>
            <td onclick="toggleStockGroup('${r.code}')" style="cursor:pointer; text-align:center;">${fmtDate(r.soonest_expiry)}</td>
            <td onclick="toggleStockGroup('${r.code}')" style="cursor:pointer; text-align:center;">
              <span class="days-chip days-${r.worst_status}">${r.min_days_remaining ?? '?'} วัน</span>
            </td>
            <td onclick="toggleStockGroup('${r.code}')" style="cursor:pointer; font-weight: 600; text-align:right;">${r.total_quantity.toLocaleString(undefined, {minimumFractionDigits: 0, maximumFractionDigits: 2})}</td>
            <td style="text-align:center;"></td>
          </tr>
        `;
        
        // Render batch sub-rows if expanded
        if (isExpanded) {
          r.batches.forEach(b => {
            rowsHtml += `
              <tr class="sub-row" style="background-color: var(--surface2); font-size: 11.5px; opacity: 0.95;">
                <td style="padding-left: 20px; text-align:center;"><span class="status-dot dot-${b.expiry_status}" style="width:6px; height:6px;"></span></td>
                <td style="color: var(--muted); font-size: 11px; text-align: center; padding-right: 12px;">└─</td>
                <td style="color: var(--muted); font-size: 11px; text-align: center;">—</td>
                <td colspan="3" style="color: var(--muted); font-size: 11px; font-style: italic; text-align:left;">รายละเอียด Batch</td>
                <td class="code-text" style="font-weight: 600; text-align:center;">${b.batch || '—'}</td>
                <td style="text-align:center;">${fmtDate(b.received_date)}</td>
                <td style="text-align:center;">${fmtDate(b.production_date)}</td>
                <td style="text-align:center;">
                  ${fmtDate(b.effective_expiry)}
                  ${b.extended_expiry_1_doc ? `<a href="${b.extended_expiry_1_doc}" target="_blank" title="ดูเอกสารต่ออายุครั้งที่ 1" style="margin-left:4px; text-decoration:none; font-size:11px;">📄</a>` : ''}
                  ${b.extended_expiry_2_doc ? `<a href="${b.extended_expiry_2_doc}" target="_blank" title="ดูเอกสารต่ออายุครั้งที่ 2" style="margin-left:4px; text-decoration:none; font-size:11px;">📄</a>` : ''}
                </td>
                <td style="text-align:center;"><span class="days-chip days-${b.expiry_status}" style="font-size:10px; padding: 2px 6px;">${b.days_remaining ?? '?'} วัน</span></td>
                <td style="font-weight: 600; font-family: var(--mono); text-align:right;">${(b.quantity ?? 0).toLocaleString(undefined, {minimumFractionDigits: 0, maximumFractionDigits: 2})}</td>
                <td style="text-align: center;">
                  ${localStorage.getItem('currentUserRole') === 'viewer' ? '' : `
                  <button class="btn btn-ghost btn-sm" onclick="openEdit('${b.id}')" style="padding: 2px 6px;">
                    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" width="12" height="12" style="stroke-width:2">
                      <path stroke-linecap="round" stroke-linejoin="round" d="M16.862 4.487l1.687-1.688a1.875 1.875 0 112.652 2.652L6.832 19.82a4.5 4.5 0 01-1.897 1.13l-2.685.8.8-2.685a4.5 4.5 0 011.13-1.897L16.863 4.487zm0 0L19.5 7.125" />
                    </svg>
                  </button>
                  `}
                </td>
              </tr>
            `;
          });
        }
      });
      tbody.innerHTML = rowsHtml;
    }
  }
  
  renderPagination('st', groupedData.length, state.stockPage, (p) => { state.stockPage=p; renderStockTable(); });
  
  // Make stock table resizable
  const table = document.querySelector('#page-stock table');
  if (table) {
    makeTableResizable(table, state.stockColWidths);
  }
}

// ---- MASTER ----
function filterMaster() {
  const searchEl = document.getElementById('m-search');
  const prodEl = document.getElementById('m-product');
  const monitorEl = document.getElementById('m-monitor-filter');
  if (!searchEl || !prodEl) return;

  const q = searchEl.value.toLowerCase();
  const pname = prodEl.value;
  const monFilter = monitorEl ? monitorEl.value : '';
  
  state.masterFiltered = state.allMaster.filter(r => {
    const step = r.recipe_step || r.rm_name || '';
    const text = `${r.code} ${r.rm_code || ''} ${r.name} ${step} ${r.product_name}`.toLowerCase();
    const monMatch = monFilter === '' ? true : monFilter === 'yes' ? r.is_monitored !== false : r.is_monitored === false;
    return (!q || text.includes(q)) && (!pname || getUnifiedProductName(r.product_name) === pname) && monMatch;
  });

  // Sort filtered data
  if (state.masterSortField) {
    const f = state.masterSortField;
    const asc = state.masterSortAsc ? 1 : -1;
    state.masterFiltered.sort((a, b) => {
      let valA, valB;
      
      if (f === 'recipe_step') {
        valA = a.recipe_step || a.rm_name || '';
        valB = b.recipe_step || b.rm_name || '';
      } else {
        valA = a[f];
        valB = b[f];
      }

      if (f === 'unit_per_pack') {
        const numA = valA !== null && valA !== undefined ? parseFloat(valA) : 0;
        const numB = valB !== null && valB !== undefined ? parseFloat(valB) : 0;
        return (numA - numB) * asc;
      }
      
      if (valA === undefined || valA === null) valA = '';
      if (valB === undefined || valB === null) valB = '';
      
      return String(valA).localeCompare(String(valB), 'th', { sensitivity: 'base' }) * asc;
    });
  }

  state.masterPage = 1;
  renderMasterTable();
}

function sortMaster(field) {
  if (state.masterSortField === field) {
    state.masterSortAsc = !state.masterSortAsc;
  } else {
    state.masterSortField = field;
    state.masterSortAsc = true;
  }
  filterMaster();
}

function renderMaster() { filterMaster(); }

async function toggleMonitor(id, currentVal) {
  const newVal = !currentVal;
  const { error } = await sb.from('raw_materials').update({ is_monitored: newVal }).eq('id', id);
  if (!error) {
    const item = state.allMaster.find(r => r.id === id);
    if (item) item.is_monitored = newVal;
    renderMasterTable();
    showToast(newVal ? 'เปิดติดตามวัตถุดิบแล้ว' : 'ปิดติดตามวัตถุดิบแล้ว', 'success');
  } else {
    showToast('เกิดข้อผิดพลาด: ' + error.message, 'error');
  }
}

function renderMasterTable() {
  // Render headers dynamically with sorting icons
  const thead = document.querySelector('#page-master table thead');
  if (thead) {
    const cols = [
      { label: 'Code', field: 'code', center: true },
      { label: 'RM Code', field: 'rm_code', center: true },
      { label: 'ProductName (สูตร)', field: 'product_name', center: false },
      { label: 'โรงงาน (Factory)', field: 'factories', center: true },
      { label: 'ชื่อวัตถุดิบ (Name)', field: 'name', center: false },
      { label: 'ขั้นตอน (Step)', field: 'recipe_step', center: true },
      { label: 'Unit', field: 'unit', center: true },
      { label: 'Package', field: 'unit_package', center: true },
      { label: 'Unit/Pack', field: 'unit_per_pack', center: true },
      { label: 'Supplier', field: 'supplier', center: true },
      { label: 'Remark', field: 'remark', center: true },
      { label: '', field: null, center: true }
    ];
    thead.innerHTML = `<tr>${cols.map(c => {
      if (!c.field) return '<th style="text-align:center;"></th>';
      const isSorted = state.masterSortField === c.field;
      const icon = isSorted ? (state.masterSortAsc ? ' ▲' : ' ▼') : ' ↕';
      const alignStyle = c.center ? 'text-align:center;' : '';
      const savedWidth = state.masterColWidths[c.field];
      const widthStyle = savedWidth ? `width: ${savedWidth}px;` : '';
      return `<th data-field="${c.field}" onclick="sortMaster('${c.field}')" class="sortable-th ${isSorted ? 'sorted-active' : ''}" style="cursor:pointer; user-select:none; position:relative; ${alignStyle} ${widthStyle}">
        <div class="th-content" style="display:inline-flex; align-items:center; gap:4px; justify-content:${c.center ? 'center' : 'flex-start'}; width:100%;">
          ${c.label} <span class="sort-icon" style="font-size:10px; color:${isSorted ? 'var(--primary)' : 'var(--muted)'};">${icon}</span>
        </div>
      </th>`;
    }).join('')}</tr>`;
  }

  const data = state.masterFiltered;
  const start = (state.masterPage-1) * state.perPage;
  const page = data.slice(start, start+state.perPage);
  
  const countEl = document.getElementById('m-count');
  if (countEl) countEl.textContent = `${data.length} รายการ`;
  
  const tbody = document.getElementById('master-tbody');
  if (tbody) {
    tbody.innerHTML = page.length
      ? page.map(r => {
          const monitored = r.is_monitored !== false;
          const toggleIcon = monitored
            ? `<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" width="14" height="14" style="stroke-width:2;color:var(--green)"><path stroke-linecap="round" stroke-linejoin="round" d="M2.036 12.322a1.012 1.012 0 010-.639C3.423 7.51 7.36 4.5 12 4.5c4.638 0 8.573 3.007 9.963 7.178.07.207.07.431 0 .639C20.577 16.49 16.64 19.5 12 19.5c-4.638 0-8.573-3.007-9.964-7.178z" /><path stroke-linecap="round" stroke-linejoin="round" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" /></svg>`
            : `<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" width="14" height="14" style="stroke-width:2;color:var(--muted)"><path stroke-linecap="round" stroke-linejoin="round" d="M3.98 8.223A10.477 10.477 0 001.934 12C3.226 16.338 7.244 19.5 12 19.5c.993 0 1.953-.138 2.863-.395M6.228 6.228A10.45 10.45 0 0112 4.5c4.756 0 8.773 3.162 10.065 7.498a10.523 10.523 0 01-4.293 5.774M6.228 6.228L3 3m3.228 3.228l3.65 3.65m7.894 7.894L21 21m-3.228-3.228l-3.65-3.65m0 0a3 3 0 10-4.243-4.243m4.242 4.242L9.88 9.88" /></svg>`;
          
          let factoriesDisplay = 'ทุกโรงงาน';
          if (r.factories && r.factories !== 'ALL') {
            factoriesDisplay = r.factories.split(',').join(', ');
          }
          
          return `<tr${monitored ? '' : ' style="opacity:0.45"'}>
            <td title="${r.code}" class="code-text" style="text-align:center;">${r.code}</td>
            <td title="${r.rm_code || ''}" class="code-text" style="text-align:center;">${r.rm_code || '—'}</td>
            <td title="${r.product_name || ''}">${r.product_name || '—'}</td>
            <td style="text-align:center;">
              <span class="badge badge-date" style="font-size:11px; padding: 2px 6px;">
                ${factoriesDisplay}
              </span>
            </td>
            <td title="${r.name || ''}">${r.name || '—'}</td>
            <td title="${r.recipe_step || r.rm_name || ''}" style="color:var(--muted); text-align:center;">${r.recipe_step || r.rm_name || '—'}</td>
            <td title="${r.unit || ''}" style="text-align:center;">${r.unit || '—'}</td>
            <td title="${r.unit_package || ''}" style="text-align:center;">${r.unit_package || '—'}</td>
            <td style="text-align:center;font-family:var(--mono)">${r.unit_per_pack ?? '—'}</td>
            <td title="${r.supplier || ''}" style="text-align:center;">${r.supplier || '—'}</td>
            <td title="${r.remark || ''}" style="color:var(--muted); font-size:11px; text-align:center;">${r.remark || '—'}</td>
            <td style="text-align:center; white-space:nowrap;">
              <button class="btn btn-ghost btn-sm" title="${monitored ? 'ปิดติดตาม' : 'เปิดติดตาม'}" onclick="toggleMonitor('${r.id}', ${monitored})" style="margin-right:2px">${toggleIcon}</button>
              <button class="btn btn-ghost btn-sm" onclick="openEditMaster('${r.id}')"><svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" width="14" height="14" style="stroke-width:2"><path stroke-linecap="round" stroke-linejoin="round" d="M16.862 4.487l1.687-1.688a1.875 1.875 0 112.652 2.652L6.832 19.82a4.5 4.5 0 01-1.897 1.13l-2.685.8.8-2.685a4.5 4.5 0 011.13-1.897L16.863 4.487zm0 0L19.5 7.125" /></svg></button>
            </td>
          </tr>`;
        }).join('')
      : '<tr><td colspan="11" class="empty"><div class="empty-icon">🗂️</div>ไม่พบข้อมูล</td></tr>';
  }
  
  renderPagination('m', data.length, state.masterPage, (p) => { state.masterPage=p; renderMasterTable(); });

  // Make master table resizable
  const table = document.querySelector('#page-master table');
  if (table) {
    makeTableResizable(table);
  }
}

function makeTableResizable(table, widthsObj) {
  const headers = table.querySelectorAll('th[data-field]');
  const widths = widthsObj || state.masterColWidths;
  
  // Set table layout to fixed if we have saved widths, so it respects them
  const hasSavedWidths = Object.keys(widths).length > 0;
  if (hasSavedWidths) {
    table.style.tableLayout = 'fixed';
  } else {
    table.style.tableLayout = 'auto';
  }
  
  headers.forEach((th) => {
    const field = th.dataset.field;
    if (!field) return;
    
    // Create resizer element
    const resizer = document.createElement('div');
    resizer.classList.add('table-resizer');
    
    // Prevent sorting click when clicking/mousedown on resizer
    resizer.addEventListener('click', (e) => e.stopPropagation());
    resizer.addEventListener('mousedown', (e) => {
      e.stopPropagation();
      e.preventDefault();
      
      const startX = e.pageX;
      const startWidth = th.offsetWidth;
      
      // Before starting to resize, set explicit widths in pixels on all ths
      if (widths === state.deductColWidths) {
        const allDeductTables = document.querySelectorAll('#deduct-queue-container table');
        allDeductTables.forEach(t => {
          t.style.tableLayout = 'fixed';
          t.querySelectorAll('th[data-field]').forEach(h => {
            h.style.width = h.offsetWidth + 'px';
          });
        });
      } else {
        headers.forEach(h => {
          h.style.width = h.offsetWidth + 'px';
        });
        table.style.tableLayout = 'fixed';
      }
      
      // Add visual active state/border during drag
      resizer.style.borderRight = '2px solid var(--accent)';
      
      function onMouseMove(e) {
        const width = startWidth + (e.pageX - startX);
        if (width > 50) {
          if (widths === state.deductColWidths) {
            document.querySelectorAll(`#deduct-queue-container th[data-field="${field}"]`).forEach(otherTh => {
              otherTh.style.width = width + 'px';
            });
          } else {
            th.style.width = width + 'px';
          }
          widths[field] = width;
        }
      }
      
      function onMouseUp() {
        document.removeEventListener('mousemove', onMouseMove);
        document.removeEventListener('mouseup', onMouseUp);
        document.body.style.cursor = '';
        resizer.style.borderRight = '';
      }
      
      document.addEventListener('mousemove', onMouseMove);
      document.addEventListener('mouseup', onMouseUp);
      document.body.style.cursor = 'col-resize';
    });
    
    th.appendChild(resizer);
  });
}

// ---- AGING REPORT ----
async function loadAging() {
  const selectEl = document.getElementById('r-date');
  if (!selectEl) return;
  const checkDate = selectEl.value;
  const tbody = document.getElementById('ag-tbody');
  if (tbody) tbody.innerHTML = '<tr><td colspan="9" class="loading"><div class="spinner"></div> กำลังโหลด...</td></tr>';
  
  let query = sb.from('stock_with_expiry').select('*').gt('quantity', 0);
  
  if (checkDate) query = query.eq('check_date', checkDate);
  else query = query.order('check_date', { ascending: false }).limit(500);
  
  const { data, error } = await query;
  if (error) {
    console.error('Error loading aging data:', error);
    showToast('เกิดข้อผิดพลาดในการดึงข้อมูล: ' + error.message, 'error');
    if (tbody) tbody.innerHTML = '<tr><td colspan="9" class="empty">เกิดข้อผิดพลาดในการดึงข้อมูล</td></tr>';
    return;
  }

  // Deduplicate by stock entry ID to ensure physical batches are only counted once
  const uniqueData = [];
  const seenIds = new Set();
  (data || []).forEach(r => {
    if (!seenIds.has(r.id)) {
      seenIds.add(r.id);
      uniqueData.push(r);
    }
  });
  
  const rows = uniqueData.map(r => {
    const enriched = enrichStock(r);
    const receivedDays = r.received_date ? Math.round((new Date() - new Date(r.received_date)) / 86400000) : null;
    let aging = '—';
    if (receivedDays !== null) {
      if (receivedDays < 90) aging = '< 3M';
      else if (receivedDays < 180) aging = '3–6M';
      else aging = '> 6M';
    }
    const supplier = r.supplier || '—';
    const unitPerPack = r.unit_per_pack || 0;
    return { ...enriched, aging, receivedDays, supplier, unitPerPack };
  });
  
  const lt3 = rows.filter(r => r.aging === '< 3M').length;
  const m36 = rows.filter(r => r.aging === '3–6M').length;
  const gt6 = rows.filter(r => r.aging === '> 6M').length;
  
  const lt3El = document.getElementById('ag-lt3');
  if (lt3El) lt3El.textContent = lt3;
  const m36El = document.getElementById('ag-36');
  if (m36El) m36El.textContent = m36;
  const gt6El = document.getElementById('ag-gt6');
  if (gt6El) gt6El.textContent = gt6;
  
  state.agingData = rows;
  
  const sorted = [...rows].sort((a,b) => (a.days_remaining??9999)-(b.days_remaining??9999));
  
  if (tbody) {
    tbody.innerHTML = sorted.length
      ? sorted.map(r => `<tr class="row-${r.expiry_status}">
          <td class="code-text">${r.code}</td>
          <td class="wrap">${r.rm_full_name || r.new_description || r.description || '—'}</td>
          <td style="font-size:12px;">${r.supplier}</td>
          <td class="code-text">${r.batch || '—'}</td>
          <td>${fmtDate(r.effective_expiry)}</td>
          <td><span class="days-chip days-${r.expiry_status}">${r.days_remaining ?? '?'} วัน</span></td>
          <td><span class="badge ${agingBadge(r.aging)}">${r.aging}</span></td>
          <td style="text-align:right">${r.quantity ?? '—'}</td>
          <td style="text-align:right;font-family:var(--mono)">${r.unitPerPack > 0 ? (r.quantity * r.unitPerPack).toFixed(0) : '—'}</td>
        </tr>`).join('')
      : '<tr><td colspan="9" class="empty">ไม่พบข้อมูล</td></tr>';
  }
}

function agingBadge(a) {
  if (a === '< 3M') return 'badge-ok';
  if (a === '3–6M') return 'badge-warn';
  if (a === '> 6M') return 'badge-expired';
  return 'badge-date';
}

function onCodeChange() {
  const val = document.getElementById('e-code').value.trim();
  const descEl = document.getElementById('e-desc');
  const supplierEl = document.getElementById('e-supplier');
  
  if (!val) {
    if (descEl) descEl.value = '';
    if (supplierEl) supplierEl.value = '';
    return;
  }
  
  // Find all master rows matching this code
  const matches = state.allMaster.filter(item => item.code === val);
  
  if (matches.length > 0) {
    if (descEl) descEl.value = matches[0].name || '';
    if (supplierEl) supplierEl.value = matches[0].supplier || '—';
  } else {
    if (descEl) descEl.value = '';
    if (supplierEl) supplierEl.value = '';
  }
}

async function saveEntry() {
  const code = document.getElementById('e-code').value;
  const batch = document.getElementById('e-batch').value.trim();
  const qty = document.getElementById('e-qty').value;
  const expiry = document.getElementById('e-expiry').value;
  
  if (!code || !batch || !expiry) {
    showToast('กรุณากรอกข้อมูลที่จำเป็น: วัตถุดิบ, Batch, วันหมดอายุ', 'error');
    return;
  }

  // ตรวจสอบว่า SAP Code มีอยู่ใน Master หรือไม่ — ถ้าไม่มีจะไม่แสดงบน Dashboard/Stock
  const codeInMaster = state.allMaster.some(m => m.code === code);
  if (!codeInMaster) {
    const proceed = confirm(`⚠️ SAP Code "${code}" ไม่พบใน Master วัตถุดิบ\n\nรายการนี้จะถูกบันทึกในฐานข้อมูลแต่จะ ไม่แสดง บน Dashboard, Stock, และหน้าวันหมดอายุ จนกว่าจะเพิ่ม Code นี้ใน Master วัตถุดิบ\n\nต้องการบันทึกต่อหรือไม่?`);
    if (!proceed) return;
  }
  
  const saveBtn = document.getElementById('e-save-btn');
  if (saveBtn) {
    saveBtn.disabled = true;
    saveBtn.dataset.originalHtml = saveBtn.innerHTML;
    saveBtn.innerHTML = `<div class="spinner" style="width:12px; height:12px; border-width:1.5px; border-top-color:#fff; border-right-color:transparent; border-bottom-color:transparent; border-left-color:transparent; margin-right:4px;"></div> กำลังบันทึก...`;
  }
  
  try {
    // Auto-lookup product_group from master data matching this code
    const matches = state.allMaster.filter(item => item.code === code);
    const grp = matches.length > 0 ? (matches[0].product_name || 'General') : 'General';
    
    const desc = document.getElementById('e-desc').value || '';
    
    const { error } = await sb.from('stock_entries').insert({
      product_group: grp,
      code: code,
      description: desc,
      new_description: desc,
      batch: batch,
      quantity: qty ? parseFloat(qty) : null,
      received_date: document.getElementById('e-received').value || null,
      production_date: document.getElementById('e-prod').value || null,
      expiry_date: expiry,
      extended_expiry_1: document.getElementById('e-ext1') ? (document.getElementById('e-ext1').value || null) : null,
      extended_expiry_2: document.getElementById('e-ext2') ? (document.getElementById('e-ext2').value || null) : null,
      shelf_life: document.getElementById('e-shelf').value || null,
      check_date: document.getElementById('e-check').value,
      notes: document.getElementById('e-notes').value || null,
    });
    
    if (error) {
      showToast(error.message, 'error');
      return;
    }
    
    showToast('บันทึกสำเร็จ!', 'success');
    resetEntryForm();
    await loadStock();
    loadRecent();
  } catch (err) {
    showToast(`เกิดข้อผิดพลาด: ${err.message}`, 'error');
  } finally {
    if (saveBtn) {
      saveBtn.disabled = false;
      saveBtn.innerHTML = saveBtn.dataset.originalHtml;
    }
  }
}

function resetEntryForm() {
  ['e-code', 'e-desc', 'e-supplier', 'e-batch', 'e-qty', 'e-prod', 'e-expiry', 'e-notes', 'e-shelf'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      if (el._flatpickr) el._flatpickr.clear();
      else el.value = '';
    }
  });
  const eReceived = document.getElementById('e-received');
  if (eReceived) {
    const today = new Date().toISOString().split('T')[0];
    if (eReceived._flatpickr) eReceived._flatpickr.setDate(today);
    else eReceived.value = today;
  }
  const eCheck = document.getElementById('e-check');
  if (eCheck) {
    const today = new Date().toISOString().split('T')[0];
    if (eCheck._flatpickr) eCheck._flatpickr.setDate(today);
    else eCheck.value = today;
  }
}

async function loadRecent() {
  const { data } = await sb.from('stock_entries').select('*').order('created_at', {ascending:false}).limit(10);
  const tbody = document.getElementById('recent-tbody');
  if (!tbody) return;
  const rows = (data || []).map(enrichStock);
  tbody.innerHTML = rows.length
    ? rows.map(r => `<tr class="row-${r.expiry_status}">
        <td class="code-text">${r.code}</td>
        <td class="wrap">${r.rm_full_name || r.new_description || r.description || '—'}</td>
        <td class="code-text">${r.batch || '—'}</td>
        <td>${r.quantity ?? '—'}</td>
        <td>${fmtDate(r.effective_expiry)}</td>
        <td style="font-size:11px;color:var(--muted)">${fmtDate(r.check_date)}</td>
        <td><button class="btn btn-ghost btn-sm" onclick="openEdit('${r.id}')"><svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" width="14" height="14" style="stroke-width:2"><path stroke-linecap="round" stroke-linejoin="round" d="M16.862 4.487l1.687-1.688a1.875 1.875 0 112.652 2.652L6.832 19.82a4.5 4.5 0 01-1.897 1.13l-2.685.8.8-2.685a4.5 4.5 0 011.13-1.897L16.863 4.487zm0 0L19.5 7.125" /></svg></button></td>
      </tr>`).join('')
    : '<tr><td colspan="7" class="empty">ยังไม่มีรายการ</td></tr>';
}

// ---- EDIT MODAL ----
function openEdit(id) {
  if (localStorage.getItem('currentUserRole') === 'viewer') {
    showToast('สิทธิ์ของคุณไม่สามารถแก้ไขข้อมูลได้', 'error');
    return;
  }
  const row = state.allStock.find(r => r.id === id);
  if (!row) return;
  state.editId = id;
  state.editExt1Doc = row.extended_expiry_1_doc || null;
  state.editExt2Doc = row.extended_expiry_2_doc || null;
  
  const formEl = document.getElementById('modal-form');
  const hasExt = !!(row.extended_expiry_1 || row.extended_expiry_2);
  if (formEl) {
    formEl.innerHTML = `
      <div class="form-group"><label class="form-label">Code</label>
        <input type="text" id="m-code" value="${row.code||''}" readonly></div>
      <div class="form-group"><label class="form-label">Batch</label>
        <input type="text" id="m-batch" value="${row.batch||''}"></div>
      <div class="form-group"><label class="form-label">จำนวน</label>
        <input type="number" id="m-qty" value="${row.quantity||''}" step="0.1"></div>
      <div class="form-group"><label class="form-label">วันรับสินค้า</label>
        <input type="date" id="m-received" value="${row.received_date||''}"></div>
      <div class="form-group"><label class="form-label">วันผลิต</label>
        <input type="date" id="m-prod" value="${row.production_date||''}"></div>
      <div class="form-group"><label class="form-label">วันหมดอายุ</label>
        <input type="date" id="m-expiry" value="${row.expiry_date||''}"></div>
      <div class="form-group full" style="display: flex; align-items: center; gap: 8px; flex-direction: row; margin-top: 8px;">
        <input type="checkbox" id="m-has-ext" ${hasExt ? 'checked' : ''} onchange="toggleModalExtensions()">
        <label class="form-label" for="m-has-ext" style="margin: 0; cursor: pointer;">มีต่ออายุวัตถุดิบ (Renewal)</label>
      </div>
      <div id="m-ext-container" class="form-group full" style="display: ${hasExt ? 'grid' : 'none'}; grid-template-columns: 1fr 1fr; gap: 16px; width: 100%; padding: 0; margin: 0;">
        <div class="form-group" style="margin:0;">
          <label class="form-label">การต่ออายุครั้งที่ 1</label>
          <input type="date" id="m-ext1" value="${row.extended_expiry_1||''}">
          <div style="margin-top: 6px;">
            <label class="form-label" style="font-size: 11px; display: block; margin-bottom: 2px;">เอกสารการต่ออายุ ครั้งที่ 1</label>
            <div style="display: flex; gap: 6px; align-items: center;">
              <button class="btn btn-ghost btn-sm" onclick="document.getElementById('m-ext1-file').click(); return false;" style="padding: 3px 8px; font-size:11px;">
                แนบไฟล์
              </button>
              <input type="file" id="m-ext1-file" style="display: none;" onchange="handleModalFileChange(1)">
              <span id="m-ext1-filename" style="font-size: 11px; color: var(--muted); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 140px;">
                ${row.extended_expiry_1_doc ? `<a href="${row.extended_expiry_1_doc}" target="_blank" style="color: var(--accent); font-weight: 600;">ดูเอกสาร 📄</a>` : 'ไม่มีไฟล์แนบ'}
              </span>
            </div>
          </div>
        </div>
        <div class="form-group" style="margin:0;">
          <label class="form-label">การต่ออายุครั้งที่ 2</label>
          <input type="date" id="m-ext2" value="${row.extended_expiry_2||''}">
          <div style="margin-top: 6px;">
            <label class="form-label" style="font-size: 11px; display: block; margin-bottom: 2px;">เอกสารการต่ออายุ ครั้งที่ 2</label>
            <div style="display: flex; gap: 6px; align-items: center;">
              <button class="btn btn-ghost btn-sm" onclick="document.getElementById('m-ext2-file').click(); return false;" style="padding: 3px 8px; font-size:11px;">
                แนบไฟล์
              </button>
              <input type="file" id="m-ext2-file" style="display: none;" onchange="handleModalFileChange(2)">
              <span id="m-ext2-filename" style="font-size: 11px; color: var(--muted); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 140px;">
                ${row.extended_expiry_2_doc ? `<a href="${row.extended_expiry_2_doc}" target="_blank" style="color: var(--accent); font-weight: 600;">ดูเอกสาร 📄</a>` : 'ไม่มีไฟล์แนบ'}
              </span>
            </div>
          </div>
        </div>
      </div>
      <div class="form-group full"><label class="form-label">หมายเหตุ</label>
        <textarea id="m-notes" rows="2">${row.notes||''}</textarea></div>
    `;
  }
  const modalEl = document.getElementById('edit-modal');
  if (modalEl) {
    modalEl.classList.add('open');
    if (typeof initFlatpickr === 'function') initFlatpickr(modalEl);
  }
}

function handleModalFileChange(num) {
  const fileInput = document.getElementById(`m-ext${num}-file`);
  const filenameSpan = document.getElementById(`m-ext${num}-filename`);
  if (fileInput && fileInput.files && fileInput.files[0]) {
    const file = fileInput.files[0];
    const reader = new FileReader();
    reader.onload = function(e) {
      state[`editExt${num}Doc`] = e.target.result; // base64 string
      filenameSpan.innerHTML = `<span style="color: var(--green); font-weight:600;">เลือกแล้ว (${file.name.slice(0, 10)}...)</span>`;
    };
    reader.readAsDataURL(file);
  }
}

function toggleModalExtensions() {
  const checkbox = document.getElementById('m-has-ext');
  const container = document.getElementById('m-ext-container');
  if (checkbox && container) {
    container.style.display = checkbox.checked ? 'grid' : 'none';
  }
}

function closeModal() {
  const modalEl = document.getElementById('edit-modal');
  if (modalEl) modalEl.classList.remove('open');
  state.editId = null;
  state.editExt1Doc = null;
  state.editExt2Doc = null;
}

async function updateEntry() {
  if (!state.editId) return;
  const hasExt = document.getElementById('m-has-ext') ? document.getElementById('m-has-ext').checked : false;
  const { error } = await sb.from('stock_entries').update({
    batch: document.getElementById('m-batch').value,
    quantity: parseFloat(document.getElementById('m-qty').value) || null,
    received_date: document.getElementById('m-received').value || null,
    production_date: document.getElementById('m-prod').value || null,
    expiry_date: document.getElementById('m-expiry').value || null,
    extended_expiry_1: hasExt ? (document.getElementById('m-ext1').value || null) : null,
    extended_expiry_2: hasExt ? (document.getElementById('m-ext2').value || null) : null,
    extended_expiry_1_doc: hasExt ? state.editExt1Doc : null,
    extended_expiry_2_doc: hasExt ? state.editExt2Doc : null,
    notes: document.getElementById('m-notes').value || null,
  }).eq('id', state.editId);
  
  if (!error) {
    closeModal();
    showToast('แก้ไขข้อมูลเรียบร้อยแล้ว!', 'success');
    await loadStock();
    const pageId = document.body.dataset.page;
    renderCurrentPage(pageId);
  } else {
    showToast('เกิดข้อผิดพลาด: ' + error.message, 'error');
  }
}

async function deleteEntry() {
  if (!state.editId || !confirm('ยืนยันการลบรายการนี้?')) return;
  const { error } = await sb.from('stock_entries').delete().eq('id', state.editId);
  if (error) {
    showToast('เกิดข้อผิดพลาดในการลบ: ' + error.message, 'error');
    return;
  }
  closeModal();
  showToast('ลบรายการเรียบร้อยแล้ว', 'success');
  await loadStock();
  const pageId = document.body.dataset.page;
  renderCurrentPage(pageId);
}

// ---- PAGINATION ----
function renderPagination(prefix, total, current, onClick) {
  const pages = Math.ceil(total / state.perPage);
  const info = document.getElementById(`${prefix}-page-info`);
  const btns = document.getElementById(`${prefix}-page-btns`);
  if (!info || !btns) return;
  
  const start = (current-1)*state.perPage+1;
  const end = Math.min(current*state.perPage, total);
  info.textContent = total ? `แสดง ${start}–${end} จาก ${total} รายการ` : '0 รายการ';
  
  let html = '';
  const range = 3;
  for (let p = 1; p <= pages; p++) {
    if (p === 1 || p === pages || (p >= current-range && p <= current+range)) {
      html += `<button class="page-btn${p===current?' active':''}" onclick="(${onClick.toString()})(${p})">${p}</button>`;
    } else if (p === current-range-1 || p === current+range+1) {
      html += `<span style="color:var(--muted);padding:4px">…</span>`;
    }
  }
  btns.innerHTML = html;
}

// ---- HELPERS ----
function getLatestStock() {
  const map = new Map();
  // Secondary sort by created_at เพื่อให้ deterministic เมื่อ check_date เดียวกัน
  const sorted = [...state.allStock].sort((a,b) => {
    const dateDiff = (b.check_date||'').localeCompare(a.check_date||'');
    if (dateDiff !== 0) return dateDiff;
    return (b.created_at||'').localeCompare(a.created_at||'');
  });
  sorted.forEach(r => {
    const key = `${r.code}__${r.batch}`;
    if (!map.has(key)) map.set(key, r);
  });
  // Only return items whose raw_material is_monitored (null = unset = show, true = show, false = hide)
  return [...map.values()].filter(r => (r.quantity ?? 0) > 0 && r.is_monitored !== false);
}

function fmtDate(d) {
  if (!d) return '—';
  try {
    const dt = new Date(d);
    if (isNaN(dt.getTime())) return d;
    // ใช้ UTC methods เพื่อป้องกัน timezone offset ทำให้วันที่คลาดเคลื่อน
    const day = String(dt.getUTCDate()).padStart(2, '0');
    const month = String(dt.getUTCMonth() + 1).padStart(2, '0');
    const year = dt.getUTCFullYear();
    return `${day}/${month}/${year}`;
  } catch { return d; }
}

// Map english state to thai display label
function statusLabel(s) {
  const m = { expired:'หมดแล้ว', critical:'≤30วัน', warning:'≤60วัน', ok:'ปกติ' };
  return m[s] || s;
}

function exportCSV(type) {
  let data, filename;
  if (type === 'expiry') {
    data = state.expiryFiltered;
    filename = 'expiry-monitor.csv';
  } else if (type === 'aging') {
    data = state.agingData;
    filename = 'aging-report.csv';
  } else {
    data = state.stockFiltered;
    filename = 'stock-all.csv';
  }
  
  if (!data.length) return;
  const keys = Object.keys(data[0]).filter(k => !['id','created_at','updated_at','raw_materials'].includes(k));
  const csv = [keys.join(','), ...data.map(r =>
    keys.map(k => `"${(r[k]??'').toString().replace(/"/g,'""')}"`).join(',')
  )].join('\n');
  
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob(['\ufeff'+csv], {type:'text/csv;charset=utf-8'}));
  a.download = filename;
  a.click();
}

function exportMasterExcel() {
  const data = state.masterFiltered;
  if (!data.length) {
    showToast('ไม่มีข้อมูลที่จะ Export', 'error');
    return;
  }
  
  // Create HTML table structure representing the excel sheet with gridlines
  let html = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
    <head>
      <meta charset="utf-8">
      <!--[if gte mso 9]>
      <xml>
        <x:ExcelWorkbook>
          <x:ExcelWorksheets>
            <x:ExcelWorksheet>
              <x:Name>Master Raw Materials</x:Name>
              <x:WorksheetOptions>
                <x:DisplayGridlines/>
              </x:WorksheetOptions>
            </x:ExcelWorksheet>
          </x:ExcelWorksheets>
        </x:ExcelWorkbook>
      </xml>
      <![endif]-->
      <style>
        table { border-collapse: collapse; }
        th { background-color: #4F46E5; color: #ffffff; font-weight: bold; }
        th, td { border: 0.5pt solid #cccccc; padding: 5px 10px; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 10pt; }
        .code { mso-number-format: "\\@"; } /* Force text format to prevent leading zero removal */
      </style>
    </head>
    <body>
      <table>
        <thead>
          <tr>
            <th>Code</th>
            <th>ProductName (สูตร)</th>
            <th>ชื่อวัตถุดิบ (Name)</th>
            <th>ขั้นตอน (Step)</th>
            <th>Unit</th>
            <th>Unit/Pack</th>
            <th>Supplier</th>
            <th>Remark</th>
            <th>ติดตาม (Monitor)</th>
          </tr>
        </thead>
        <tbody>
  `;
  
  data.forEach(r => {
    const monitored = r.is_monitored !== false ? 'ติดตาม' : 'ไม่ติดตาม';
    html += `
      <tr>
        <td class="code">${r.code || ''}</td>
        <td>${r.product_name || ''}</td>
        <td>${r.name || ''}</td>
        <td>${r.recipe_step || r.rm_name || ''}</td>
        <td>${r.unit || ''}</td>
        <td style="text-align:right;">${r.unit_per_pack ?? ''}</td>
        <td>${r.supplier || ''}</td>
        <td>${r.remark || ''}</td>
        <td>${monitored}</td>
      </tr>
    `;
  });
  
  html += `
        </tbody>
      </table>
    </body>
    </html>
  `;
  
  const blob = new Blob(['\ufeff' + html], { type: 'application/vnd.ms-excel;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `master-raw-materials_${new Date().toISOString().split('T')[0]}.xls`;
  a.click();
}

function toggleSidebar() {
  const sidebar = document.querySelector('.sidebar');
  if (!sidebar) return;
  
  if (window.innerWidth > 768) {
    const isCollapsed = document.body.classList.toggle('sidebar-collapsed');
    localStorage.setItem('sidebar-collapsed', isCollapsed ? 'true' : 'false');
  } else {
    sidebar.classList.toggle('open');
  }
}

function setExpiryShortcut(months) {
  const prodVal = document.getElementById('e-prod').value;
  const receivedVal = document.getElementById('e-received').value;
  
  let baseDateStr = prodVal || receivedVal;
  let baseDate = baseDateStr ? new Date(baseDateStr) : new Date();
  
  if (isNaN(baseDate.getTime())) {
    baseDate = new Date();
  }
  
  baseDate.setMonth(baseDate.getMonth() + months);
  
  const yyyy = baseDate.getFullYear();
  const mm = String(baseDate.getMonth() + 1).padStart(2, '0');
  const dd = String(baseDate.getDate()).padStart(2, '0');
  
  const expiryEl = document.getElementById('e-expiry');
  if (expiryEl) {
    const formatted = `${yyyy}-${mm}-${dd}`;
    if (expiryEl._flatpickr) expiryEl._flatpickr.setDate(formatted);
    else expiryEl.value = formatted;
  }
  
  const shelfEl = document.getElementById('e-shelf');
  if (shelfEl) {
    if (months === 6) {
      shelfEl.value = '6 เดือน';
    } else if (months === 12) {
      shelfEl.value = '1 ปี';
    } else if (months === 24) {
      shelfEl.value = '2 ปี';
    } else {
      shelfEl.value = `${months} เดือน`;
    }
  }
}

function initFlatpickr(container = document) {
  if (typeof flatpickr === 'undefined') return;
  container.querySelectorAll('input[type="date"]').forEach(el => {
    flatpickr(el, {
      dateFormat: "Y-m-d",
      altInput: true,
      altFormat: "d/m/Y",
      allowInput: true,
    });
  });
}

function showToast(message, type = 'success') {
  let container = document.getElementById('toast-container');
  if (!container) {
    container = document.createElement('div');
    container.id = 'toast-container';
    container.style.position = 'fixed';
    container.style.top = '24px';
    container.style.right = '24px';
    container.style.zIndex = '9999';
    container.style.display = 'flex';
    container.style.flexDirection = 'column';
    container.style.gap = '10px';
    document.body.appendChild(container);
  }
  
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.style.background = type === 'success' ? 'var(--green)' : 'var(--red)';
  toast.style.color = '#fff';
  toast.style.padding = '12px 20px';
  toast.style.borderRadius = 'var(--radius)';
  toast.style.boxShadow = '0 10px 15px -3px rgba(0,0,0,0.1), 0 4px 6px -2px rgba(0,0,0,0.05)';
  toast.style.fontSize = '14px';
  toast.style.fontWeight = '500';
  toast.style.display = 'flex';
  toast.style.alignItems = 'center';
  toast.style.gap = '8px';
  toast.style.minWidth = '280px';
  toast.style.transition = 'all 0.3s ease';
  toast.style.transform = 'translateY(-20px)';
  toast.style.opacity = '0';
  
  const icon = type === 'success' ? '✅' : '⚠️';
  toast.innerHTML = `<span>${icon}</span> <span>${message}</span>`;
  
  container.appendChild(toast);
  
  setTimeout(() => {
    toast.style.transform = 'translateY(0)';
    toast.style.opacity = '1';
  }, 10);
  
  setTimeout(() => {
    toast.style.transform = 'translateY(-20px)';
    toast.style.opacity = '0';
    setTimeout(() => {
      toast.remove();
    }, 300);
  }, 3000);
}

// ---- KPI DETAILS MODAL ----
function openKpiDetails(type) {
  const latest = getLatestStock();
  let filtered = [];
  let title = '';
  
  if (type === 'total') {
    filtered = latest;
    title = 'วัตถุดิบคงคลังทั้งหมด';
  } else if (type === 'expired') {
    filtered = latest.filter(s => s.expiry_status === 'expired');
    title = 'วัตถุดิบหมดอายุ';
  } else if (type === 'critical') {
    filtered = latest.filter(s => s.expiry_status === 'critical');
    title = 'วัตถุดิบใกล้หมดอายุมาก (≤ 30 วัน)';
  } else if (type === 'warning') {
    filtered = latest.filter(s => s.expiry_status === 'warning');
    title = 'วัตถุดิบควรติดตาม (31–60 วัน)';
  } else if (type === 'ok') {
    filtered = latest.filter(s => s.expiry_status === 'ok');
    title = 'วัตถุดิบปกติ (> 60 วัน)';
  }

  const titleEl = document.getElementById('kpi-modal-title');
  if (titleEl) titleEl.textContent = `${title} (${filtered.length} รายการ)`;

  const tbody = document.getElementById('kpi-modal-tbody');
  if (tbody) {
    tbody.innerHTML = filtered.length
      ? filtered.map(r => `
          <tr class="row-${r.expiry_status}">
            <td><span class="status-dot dot-${r.expiry_status}"></span></td>
            <td class="code-text">${r.code}</td>
            <td class="wrap">${r.rm_full_name || r.new_description || r.description || '—'}</td>
            <td>${r.product_name || '—'}</td>
            <td class="code-text">${r.batch || '—'}</td>
            <td><span class="days-chip days-${r.expiry_status}">${r.days_remaining ?? '?'} วัน</span></td>
            <td>${r.quantity ?? '—'}</td>
          </tr>
        `).join('')
      : '<tr><td colspan="7" class="empty">ไม่มีรายการในหมวดหมู่นี้</td></tr>';
  }

  const modal = document.getElementById('kpi-modal');
  if (modal) modal.classList.add('open');
}

function closeKpiModal() {
  const modal = document.getElementById('kpi-modal');
  if (modal) modal.classList.remove('open');
}

// ---- EDIT MASTER MODAL ----
function toggleEditMasterFactories(el) {
  if (el.checked) {
    document.querySelectorAll('.m-fac-spec').forEach(cb => cb.checked = false);
  }
}

function toggleEditMasterSpecFactories(el) {
  if (el.checked) {
    const allCb = document.getElementById('m-fac-all');
    if (allCb) allCb.checked = false;
  }
}

function openEditMaster(id) {
  const row = state.allMaster.find(r => r.id === id);
  if (!row) return;
  state.editMasterId = id;
  
  const formEl = document.getElementById('modal-form');
  if (formEl) {
    const stepVal = row.recipe_step || row.rm_name || '';
    const monitored = row.is_monitored !== false;
    
    const facVal = row.factories || 'ALL';
    const isAll = facVal.includes('ALL');
    const isCH = facVal.includes('CH');
    const isSR = facVal.includes('SR');
    const isKR = facVal.includes('KR');
    const isNS = facVal.includes('NS');
    
    formEl.innerHTML = `
      <div class="form-group"><label class="form-label">Code (SAP Code)</label>
        <input type="text" id="m-code" value="${row.code||''}"></div>
      <div class="form-group"><label class="form-label">RM Code</label>
        <input type="text" id="m-rm-code" value="${row.rm_code||''}"></div>
      <div class="form-group"><label class="form-label">ProductName (สูตร)</label>
        <input type="text" id="m-product-name" value="${row.product_name||''}"></div>
      <div class="form-group"><label class="form-label">ชื่อวัตถุดิบ (Name)</label>
        <input type="text" id="m-name" value="${row.name||''}"></div>
      <div class="form-group"><label class="form-label">ขั้นตอน (Step)</label>
        <input type="text" id="m-recipe-step" value="${stepVal}"></div>
      <div class="form-group"><label class="form-label">Unit</label>
        <input type="text" id="m-unit" value="${row.unit||''}"></div>
      <div class="form-group"><label class="form-label">Package</label>
        <input type="text" id="m-unit-package" value="${row.unit_package||''}"></div>
      <div class="form-group"><label class="form-label">Unit/Pack</label>
        <input type="number" id="m-unit-per-pack" value="${row.unit_per_pack??'0'}" step="0.0001"></div>
      <div class="form-group"><label class="form-label">Supplier</label>
        <input type="text" id="m-supplier" value="${row.supplier||''}"></div>
      <div class="form-group full" style="display:flex;align-items:center;gap:10px;flex-direction:row;">
        <input type="checkbox" id="m-is-monitored" ${monitored ? 'checked' : ''} style="width:16px;height:16px;cursor:pointer;">
        <label class="form-label" for="m-is-monitored" style="margin:0;cursor:pointer;">ติดตามวัตถุดิบนี้ (Monitor) — ถ้าไม่ติ๊ก จะไม่แสดงใน Dashboard และรายงานต่างๆ</label>
      </div>
      
      <div class="form-group full" style="margin-top: 8px;">
        <label class="form-label" style="font-weight: 600;">โรงงานที่ใช้งานวัตถุดิบนี้ (Factory Restrictions)</label>
        <div style="display: flex; gap: 16px; flex-wrap: wrap; margin-top: 6px; background: var(--surface2); padding: 10px; border-radius: 6px; border: 1px solid var(--border);">
          <label style="cursor:pointer; display:inline-flex; align-items:center; gap:6px;">
            <input type="checkbox" id="m-fac-all" value="ALL" ${isAll ? 'checked' : ''} onchange="toggleEditMasterFactories(this)"> ทุกโรงงาน (ALL)
          </label>
          <label style="cursor:pointer; display:inline-flex; align-items:center; gap:6px;">
            <input type="checkbox" class="m-fac-spec" value="CH" ${isCH ? 'checked' : ''} onchange="toggleEditMasterSpecFactories(this)"> ชลบุรี (CH)
          </label>
          <label style="cursor:pointer; display:inline-flex; align-items:center; gap:6px;">
            <input type="checkbox" class="m-fac-spec" value="SR" ${isSR ? 'checked' : ''} onchange="toggleEditMasterSpecFactories(this)"> สุราษฎร์ธานี (SR)
          </label>
          <label style="cursor:pointer; display:inline-flex; align-items:center; gap:6px;">
            <input type="checkbox" class="m-fac-spec" value="KR" ${isKR ? 'checked' : ''} onchange="toggleEditMasterSpecFactories(this)"> โคราช (KR)
          </label>
          <label style="cursor:pointer; display:inline-flex; align-items:center; gap:6px;">
            <input type="checkbox" class="m-fac-spec" value="NS" ${isNS ? 'checked' : ''} onchange="toggleEditMasterSpecFactories(this)"> นครสวรรค์ (NS)
          </label>
        </div>
      </div>

      <div class="form-group full"><label class="form-label">หมายเหตุ (Remark)</label>
        <textarea id="m-remark" rows="2">${row.remark||''}</textarea></div>
    `;
  }
  const modalEl = document.getElementById('edit-modal');
  if (modalEl) modalEl.classList.add('open');
}

async function updateMasterEntry() {
  if (!state.editMasterId) return;
  const isMonitoredEl = document.getElementById('m-is-monitored');
  
  let factoriesStr = 'ALL';
  const allCb = document.getElementById('m-fac-all');
  if (allCb && allCb.checked) {
    factoriesStr = 'ALL';
  } else {
    const selected = [];
    document.querySelectorAll('.m-fac-spec').forEach(cb => {
      if (cb.checked) selected.push(cb.value);
    });
    factoriesStr = selected.length > 0 ? selected.join(',') : 'ALL';
  }

  const { error } = await sb.from('raw_materials').update({
    code: document.getElementById('m-code').value.trim(),
    rm_code: document.getElementById('m-rm-code').value.trim(),
    product_name: document.getElementById('m-product-name').value.trim(),
    name: document.getElementById('m-name').value.trim(),
    recipe_step: document.getElementById('m-recipe-step').value.trim(),
    unit: document.getElementById('m-unit').value.trim(),
    unit_package: document.getElementById('m-unit-package').value.trim(),
    unit_per_pack: parseFloat(document.getElementById('m-unit-per-pack').value) || 0,
    supplier: document.getElementById('m-supplier').value.trim(),
    remark: document.getElementById('m-remark').value.trim(),
    is_monitored: isMonitoredEl ? isMonitoredEl.checked : true,
    factories: factoriesStr,
  }).eq('id', state.editMasterId);
  
  if (!error) {
    closeModal();
    showToast('แก้ไขข้อมูลวัตถุดิบเรียบร้อยแล้ว!', 'success');
    await loadMaster();
    filterMaster();
  } else {
    showToast('เกิดข้อผิดพลาด: ' + error.message, 'error');
  }
}

async function deleteMasterEntry() {
  if (!state.editMasterId || !confirm('ยืนยันการลบข้อมูลวัตถุดิบตัวนี้? การลบข้อมูล Master อาจส่งผลต่อยอดคงเหลือวัตถุดิบใน Stock')) return;
  const { error } = await sb.from('raw_materials').delete().eq('id', state.editMasterId);
  if (!error) {
    closeModal();
    showToast('ลบข้อมูลเรียบร้อยแล้ว!', 'success');
    await loadMaster();
    filterMaster();
  } else {
    showToast('เกิดข้อผิดพลาด: ' + error.message, 'error');
  }
}

// ---- STOCK DEDUCTION (ระบบตัดจ่ายวัตถุดิบ) ----
let deductQueue = []; // holds objects: { id, productName, unitsToProduce, ingredients: [ { code, name, recipeStep, qtyPerUnit, reqQty, unit, remark, selectedBatch, batches: [...] } ] }

function getUnifiedProductName(dbName) {
  if (!dbName) return "";
  if (dbName === "Cola 5.3 (CL 5.3)" || dbName === "Cola 5.3 (CH,SR)") {
    return "Cola 5.3 (CL 5.3)";
  }
  if (dbName === "Cream Soda 5.4" || dbName === "Cream Soda 5.4 (SR,CH)" || 
      dbName === "Cream Soda 5.4 (KR,NS)" || dbName === "Cream Soda 5.4 (PT)" || 
      dbName === "CS 5.4 (CH,SR)") {
    return "Cream Soda 5.4";
  }
  if (dbName === "Grape Berry 5.4" || dbName === "GB 5.4 (CH,SR)") {
    return "Grape Berry 5.4";
  }
  if (dbName === "Salty Lychee 5.4 (Giv)" || dbName === "LC 5.4 (CH,SR)") {
    return "Salty Lychee 5.4 (Giv)";
  }
  if (dbName === "Lemon Lime 5.4" || dbName === "Lemon Lime 5.4 (CH,SR)") {
    return "Lemon Lime 5.4";
  }
  if (dbName === "Orange 5.4" || dbName === "Orange 5.4 (SR,CH)") {
    return "Orange 5.4";
  }
  if (dbName === "Strawberry 5.4" || dbName === "Strawberry 5.4 (CH,SR)") {
    return "Strawberry 5.4";
  }
  if (dbName === "Cola 7.3 (CLL 7.3)" || dbName === "Cola 7.3 (SR)") {
    return "Cola 7.3 (CLL 7.3)";
  }
  return dbName;
}

function getDeductIngredients(unifiedProdName, factory) {
  // 1. Get all raw materials that belong to the unified product group
  const baseIngredients = state.allMaster.filter(m => 
    getUnifiedProductName(m.product_name) === unifiedProdName && 
    m.is_monitored !== false
  );
  
  // 2. Filter ingredients based on factory restrictions
  return baseIngredients.filter(m => {
    const facVal = (m.factories || 'ALL').trim().toUpperCase();
    
    // If factories is 'ALL', it's allowed for any factory selection
    if (facVal === 'ALL' || facVal === '') {
      return true;
    }
    
    // If no specific factory is selected (General), only allow ALL or General
    if (!factory) {
      const allowedList = facVal.split(',').map(s => s.trim());
      return allowedList.includes('GENERAL');
    }
    
    // Check if the selected factory code is explicitly allowed in the comma-separated list
    const allowedList = facVal.split(',').map(s => s.trim());
    return allowedList.includes(factory.toUpperCase());
  });
}

function renderDeduct() {
  const seenUnified = new Set();
  const pnames = [];
  
  state.allMaster.forEach(m => {
    if (!m.product_name) return;
    const unified = getUnifiedProductName(m.product_name);
    if (!seenUnified.has(unified)) {
      seenUnified.add(unified);
      pnames.push(unified);
    }
  });
  
  pnames.sort((a, b) => a.localeCompare(b, 'th', { sensitivity: 'base' }));
  
  const sel = document.getElementById('d-product');
  if (sel) {
    sel.innerHTML = '<option value="">— เลือกสูตรการผลิต —</option>';
    pnames.forEach(p => {
      const opt = document.createElement('option');
      opt.value = p;
      opt.textContent = p;
      sel.appendChild(opt);
    });
  }
  updateDeductQueueUI();
}

function onDeductProductChange() {
  // Can be used for immediate reaction if needed.
}

function onDeductFactoryChange() {
  // Can be used for immediate reaction if needed.
}

function getBatchesForCode(code) {
  // Accumulate stock batch balances based on latest stock entries
  const batchMap = new Map();
  // Filter allStock entries matching this code
  const items = state.allStock.filter(s => s.code === code);
  
  // Sort by check_date descending to accumulate current quantity per batch
  const sorted = [...items].sort((a,b) => (b.check_date||'').localeCompare(a.check_date||''));
  
  sorted.forEach(s => {
    const key = s.batch || 'no-batch';
    if (!batchMap.has(key)) {
      // Exclude zero or negative stock from deduction selectables
      if (s.quantity > 0) {
        batchMap.set(key, {
          batch: s.batch,
          quantity: s.quantity,
          effective_expiry: s.effective_expiry,
          expiry_status: s.expiry_status,
          days_remaining: s.days_remaining,
          description: s.new_description || s.description || ''
        });
      }
    }
  });
  
  // Sort by effective_expiry ascending (FEFO - First Expired First Out)
  return [...batchMap.values()].sort((a, b) => {
    if (!a.effective_expiry) return 1;
    if (!b.effective_expiry) return -1;
    return a.effective_expiry.localeCompare(b.effective_expiry);
  });
}

function addDeductionGroup() {
  const prodName = document.getElementById('d-product').value;
  const unitsVal = parseInt(document.getElementById('d-unit-qty').value);
  const factory = document.getElementById('d-factory').value;
  
  if (!prodName) {
    showToast('กรุณาเลือกสูตรการผลิต', 'error');
    return;
  }
  if (!unitsVal || unitsVal < 1) {
    showToast('กรุณาระบุจำนวนยูนิตผลิตที่ถูกต้อง (>= 1)', 'error');
    return;
  }
  
  const formulaRMs = getDeductIngredients(prodName, factory);
  if (formulaRMs.length === 0) {
    showToast('ไม่พบส่วนผสมในสูตรนี้ในข้อมูล Master', 'error');
    return;
  }
  
  const ingredients = formulaRMs.map(m => {
    // Each ingredient in master is assumed to be 1 Unit/Pack as base
    // If unit_per_pack is specified, 1 pack produces unit_per_pack units.
    // Therefore, to produce unitsVal units, we need: unitsVal / unit_per_pack packs.
    const divisor = m.unit_per_pack && m.unit_per_pack > 0 ? m.unit_per_pack : 1;
    const reqQty = unitsVal / divisor;
    
    // Get list of active stock batches for this code
    const batches = getBatchesForCode(m.code);
    
    return {
      code: m.code,
      name: m.name,
      recipeStep: m.recipe_step || m.rm_name || '',
      qtyPerUnit: 1 / divisor,
      reqQty: parseFloat(reqQty.toFixed(4)),
      unit: m.unit || 'UN',
      remark: m.remark || '',
      batches: batches,
      selectedBatch: batches.length > 0 ? batches[0].batch : '', // Default to earliest expiring batch
      selected: true,
      origProductName: m.product_name
    };
  });
  
  const queueItem = {
    id: 'dq_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9),
    productName: prodName,
    factory: factory,
    unitsToProduce: unitsVal,
    ingredients: ingredients
  };
  
  deductQueue.push(queueItem);
  
  // Reset form
  document.getElementById('d-product').value = '';
  document.getElementById('d-unit-qty').value = '1';
  document.getElementById('d-factory').value = '';
  
  updateDeductQueueUI();
  showToast('เพิ่มสูตรลงรายการเรียบร้อย', 'success');
}

function removeQueueItem(id) {
  deductQueue = deductQueue.filter(item => item.id !== id);
  updateDeductQueueUI();
}

function clearDeductQueue() {
  if (deductQueue.length === 0 || confirm('ต้องการล้างรายการทั้งหมด?')) {
    deductQueue = [];
    updateDeductQueueUI();
  }
}

function updateIngredientBatch(groupId, ingredientCode, batchVal) {
  const group = deductQueue.find(g => g.id === groupId);
  if (group) {
    const ing = group.ingredients.find(i => i.code === ingredientCode);
    if (ing) {
      ing.selectedBatch = batchVal;
      updateDeductQueueUI();
    }
  }
}

function toggleSelectAllDeductGroup(groupId, checked) {
  const group = deductQueue.find(g => g.id === groupId);
  if (group) {
    group.ingredients.forEach(ing => {
      ing.selected = checked;
    });
    updateDeductQueueUI();
  }
}

function toggleSelectDeductIngredient(groupId, ingredientCode, checked) {
  const group = deductQueue.find(g => g.id === groupId);
  if (group) {
    const ing = group.ingredients.find(i => i.code === ingredientCode);
    if (ing) {
      ing.selected = checked;
      updateDeductQueueUI();
    }
  }
}

function updateDeductIngredientQty(groupId, ingredientCode, qtyVal) {
  const group = deductQueue.find(g => g.id === groupId);
  if (group) {
    const ing = group.ingredients.find(i => i.code === ingredientCode);
    if (ing) {
      const parsed = parseFloat(qtyVal);
      ing.reqQty = isNaN(parsed) ? 0 : parsed;
    }
  }
}

function updateDeductQueueUI() {
  const container = document.getElementById('deduct-queue-container');
  const submitBar = document.getElementById('deduct-submit-bar');
  const clearBtn = document.getElementById('clear-queue-btn');
  
  if (!container) return;
  
  const getColWidth = (field, defaultWidth) => {
    const saved = state.deductColWidths && state.deductColWidths[field];
    if (saved) return `width: ${saved}px;`;
    return defaultWidth ? `width: ${defaultWidth};` : '';
  };
  
  if (deductQueue.length === 0) {
    container.innerHTML = `
      <div class="empty" style="padding: 40px; text-align: center; color: var(--muted);">
        <div class="empty-icon" style="font-size: 32px; margin-bottom: 8px;">📋</div>
        ยังไม่มีรายการเบิก เลือกสูตรและระบุจำนวนเพื่อเพิ่มเข้ารายการ
      </div>
    `;
    if (submitBar) submitBar.style.display = 'none';
    if (clearBtn) clearBtn.style.display = 'none';
    return;
  }
  
  if (submitBar) submitBar.style.display = 'flex';
  if (clearBtn) clearBtn.style.display = 'block';
  
  let html = '';
  
  deductQueue.forEach(group => {
    let ingHtml = '';
    
    // Check if all ingredients in this group are selected
    const allSelected = group.ingredients.every(ing => ing.selected !== false);
    
    group.ingredients.forEach(ing => {
      // Build batch options dropdown
      let batchOptions = '';
      let selectedBatchInfo = null;
      
      if (ing.batches.length === 0) {
        batchOptions = '<option value="">⚠️ ไม่มีของในคลัง</option>';
      } else {
        ing.batches.forEach(b => {
          const expText = b.effective_expiry ? ` (EXP: ${fmtDate(b.effective_expiry)})` : '';
          const qtyText = `คงเหลือ: ${b.quantity} ${ing.unit}`;
          const isSelected = b.batch === ing.selectedBatch ? ' selected' : '';
          batchOptions += `<option value="${b.batch}"${isSelected}>Batch: ${b.batch} | ${qtyText}${expText}</option>`;
          
          if (b.batch === ing.selectedBatch) {
            selectedBatchInfo = b;
          }
        });
      }
      
      let badgeHtml = '';
      if (selectedBatchInfo) {
        badgeHtml = `<span class="days-chip days-${selectedBatchInfo.expiry_status}" style="font-size:10px; padding: 1px 6px;">${selectedBatchInfo.days_remaining} วัน</span>`;
      }
      
      ingHtml += `
        <tr style="border-bottom: 1px solid var(--border-soft); opacity: ${ing.selected !== false ? '1.0' : '0.55'}">
          <td style="padding: 8px 12px; text-align: center; width: 40px; vertical-align: middle;">
            <input type="checkbox" onchange="toggleSelectDeductIngredient('${group.id}', '${ing.code}', this.checked)" ${ing.selected !== false ? 'checked' : ''}>
          </td>
          <td class="code-text" style="padding: 8px 12px; font-size:12px;">${ing.code}</td>
          <td style="padding: 8px 12px; font-size:12px;">
            <div><strong>${ing.name || '—'}</strong></div>
            <div style="font-size: 11px; color: var(--muted)">${ing.recipeStep}</div>
          </td>
          <td style="padding: 8px 12px; font-size:12px; font-weight: 600; text-align: right; color: var(--accent);">
            <div style="display: inline-flex; align-items: center; gap: 4px; justify-content: flex-end; width: 100%;">
              <input type="number" step="any" min="0" value="${ing.reqQty}" 
                onchange="updateDeductIngredientQty('${group.id}', '${ing.code}', this.value)" 
                style="width: 70px; text-align: right; font-size: 12px; padding: 4px 6px; border-radius: 4px; border: 1px solid var(--border); background: var(--surface);" 
                ${ing.selected === false ? 'disabled' : ''}>
              <span>${ing.unit}</span>
            </div>
          </td>
          <td style="padding: 8px 12px; font-size:12px;">
            <select style="width: 100%; font-size: 12px; padding: 4px 8px; border-radius: 4px;" 
              onchange="updateIngredientBatch('${group.id}', '${ing.code}', this.value)"
              ${ing.selected === false ? 'disabled' : ''}>
              ${batchOptions}
            </select>
          </td>
          <td style="padding: 8px 12px; text-align: center;">
            ${ing.selected !== false ? badgeHtml : ''}
          </td>
        </tr>
      `;
    });
    
    html += `
      <div class="card" style="margin-bottom: 16px; border: 1px solid var(--border); background: var(--surface2); padding: 16px; position: relative;">
        <button class="btn btn-ghost btn-sm" onclick="removeQueueItem('${group.id}')" style="position: absolute; top: 12px; right: 12px; color: var(--red); padding: 4px; border: none; background: transparent;">
          <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" width="16" height="16" style="stroke-width:2">
            <path stroke-linecap="round" stroke-linejoin="round" d="M6 18L18 6M6 6l12 12" />
          </svg>
        </button>
        <div style="margin-bottom: 12px;">
          <span style="font-size: 15px; font-weight: 700; color: var(--text);">${group.productName}</span>
          <span class="badge badge-date" style="margin-left: 8px;">โรงงาน: ${group.factory ? group.factory : 'ทั่วไป (General)'}</span>
          <span class="badge badge-date" style="margin-left: 8px;">ผลิต ${group.unitsToProduce} UNIT</span>
        </div>
        <div class="table-wrapper" style="border-radius: 6px; border: 1px solid var(--border); overflow: visible;">
          <table style="width: 100%; border-collapse: collapse; background: var(--surface);">
            <thead>
              <tr style="background: var(--surface2); border-bottom: 1px solid var(--border);">
                <th style="padding: 6px 12px; font-size:11px; text-align: center; width: 40px;">
                  <input type="checkbox" onchange="toggleSelectAllDeductGroup('${group.id}', this.checked)" ${allSelected ? 'checked' : ''}>
                </th>
                <th data-field="sap_code" style="padding: 6px 12px; font-size:11px; text-align: center; position: relative; ${getColWidth('sap_code')}">
                  <div style="display: inline-flex; align-items: center; justify-content: center; width: 100%;">รหัส SAP</div>
                </th>
                <th data-field="ingredient_name" style="padding: 6px 12px; font-size:11px; text-align: center; position: relative; ${getColWidth('ingredient_name')}">
                  <div style="display: inline-flex; align-items: center; justify-content: center; width: 100%;">รายการส่วนผสม</div>
                </th>
                <th data-field="quantity" style="padding: 6px 12px; font-size:11px; text-align: center; position: relative; ${getColWidth('quantity', '120px')}">
                  <div style="display: inline-flex; align-items: center; justify-content: center; width: 100%;">ต้องใช้</div>
                </th>
                <th data-field="batch_select" style="padding: 6px 12px; font-size:11px; text-align: center; position: relative; ${getColWidth('batch_select')}">
                  <div style="display: inline-flex; align-items: center; justify-content: center; width: 100%;">เลือก Batch จ่ายคลัง</div>
                </th>
                <th data-field="days_remaining" style="padding: 6px 12px; font-size:11px; text-align: center; position: relative; ${getColWidth('days_remaining', '90px')}">
                  <div style="display: inline-flex; align-items: center; justify-content: center; width: 100%;">อายุคงเหลือ</div>
                </th>
              </tr>
            </thead>
            <tbody>
              ${ingHtml}
            </tbody>
          </table>
        </div>
      </div>
    `;
  });
  
  container.innerHTML = html;
  
  // Make all tables in deduct queue resizable
  container.querySelectorAll('table').forEach(table => {
    makeTableResizable(table, state.deductColWidths);
  });
}

async function submitDeductionQueue() {
  if (deductQueue.length === 0) return;
  
  // Verify that at least one ingredient is checked across all groups
  const hasSelected = deductQueue.some(group => group.ingredients.some(ing => ing.selected !== false));
  if (!hasSelected) {
    showToast('กรุณาเลือกวัตถุดิบอย่างน้อย 1 รายการเพื่อตัดสต็อก', 'error');
    return;
  }
  
  const submitBtn = document.getElementById('deduct-submit-btn');
  if (submitBtn) {
    submitBtn.disabled = true;
    submitBtn.dataset.originalHtml = submitBtn.innerHTML;
    submitBtn.innerHTML = `<div class="spinner" style="width:12px; height:12px; border-width:1.5px; border-top-color:#fff; border-right-color:transparent; border-bottom-color:transparent; border-left-color:transparent; margin-right:4px;"></div> กำลังบันทึกตัดสต็อก...`;
  }
  
  try {
    const todayStr = new Date().toISOString().split('T')[0];
    const insertData = [];
    
    // Validate that all items have a selected batch
    for (const group of deductQueue) {
      for (const ing of group.ingredients) {
        if (ing.selected === false) continue; // Skip unchecked ingredients
        
        if (!ing.selectedBatch) {
          throw new Error(`ไม่พบ Batch คงเหลือสำหรับวัตถุดิบ ${ing.code} (${ing.name}) ในสูตร ${group.productName}`);
        }
        
        // Find batch details to fetch description/new_description and other defaults
        const batchInfo = ing.batches.find(b => b.batch === ing.selectedBatch);
        
        const grp = ing.origProductName || 'General';
        
        insertData.push({
          product_group: grp,
          code: ing.code,
          description: batchInfo ? batchInfo.description : (ing.name || ''),
          new_description: batchInfo ? batchInfo.description : (ing.name || ''),
          batch: ing.selectedBatch,
          quantity: -ing.reqQty, // Deduction (negative quantity)
          received_date: todayStr,
          check_date: todayStr,
          notes: `เบิกผลิตสูตร ${group.productName} (${group.factory || 'ทั่วไป'}) จำนวน ${group.unitsToProduce} UNIT โดย ${localStorage.getItem('currentUserName') || 'ไม่ระบุผู้ใช้'}`
        });
      }
    }
    
    // Insert into database in batches
    for (let i = 0; i < insertData.length; i += 100) {
      const batch = insertData.slice(i, i + 100);
      const { error } = await sb.from('stock_entries').insert(batch);
      if (error) throw error;
    }
    
    showToast('เบิกตัดสต็อกวัตถุดิบสำเร็จ!', 'success');
    deductQueue = [];
    await loadStock(); // Reload updated stock levels
    renderDeduct();
    
  } catch (err) {
    showToast(`เกิดข้อผิดพลาด: ${err.message}`, 'error');
  } finally {
    if (submitBtn) {
      submitBtn.disabled = false;
      submitBtn.innerHTML = submitBtn.dataset.originalHtml;
    }
  }
}

// ---- EMAIL PREVIEW & COMPLETE SET PRODUCTION CALCULATION ----
const staticReportItems = [
  { code: '120000965', name: 'Cola compound Flavor FA-01L', formula: 'Cola 7.3 (CLL 7.3)', showComplete: true, borderColor: '#ed8936' },
  { code: '120001794', name: 'Cola Compound Flavor FA-35L @20.4', formula: 'Cola 7.3 (CLL 7.3)', showComplete: false, borderColor: '#ed8936' },
  { code: '120002038', name: 'Cola Compound Flavor FA-35L @22.50 kg', formula: 'Cola 5.3 (CL 5.3)', showComplete: false, borderColor: '#ed8936' },
  { code: '120001436', name: 'Orange Emulsion Compound FA-07L @19.635', formula: 'Orange 5.4', showComplete: true, borderColor: '#f6ad55' },
  { code: '120001882', name: 'Cream Soda Flavor FA-06L @12.00', formula: 'Cream Soda 5.4', showComplete: false, borderColor: '#48bb78' },
  { code: '120001869', name: '941849 Strawberry Flavour FA-37L @18.0', formula: 'Strawberry 5.4', showComplete: true, borderColor: '#e53e3e' },
  { code: '120001889', name: 'Lemon Lime Flavor FA-05L @13.17', formula: 'Lemon Lime 5.4', showComplete: true, borderColor: '#38a169' },
  { code: '120002205', name: 'Salty Lemonade Flavour FA-50L@5.25', formula: 'Salty Lemonade', showComplete: true, borderColor: '#2f855a' },
  { code: '120001868', name: 'Kamikaze Lime Flavor FA-34L @23.1', formula: 'Kamikaze Lime 5.4 (Giv)', showComplete: true, borderColor: '#3182ce' },
  { code: '120001627', name: 'Salty Lychee Flavor FA-27L @7.04', formula: 'Salty Lychee 5.4 (Giv)', showComplete: true, borderColor: '#d53f8c' },
  { code: '120001628', name: 'Salty lychee flavor FA-28L @4.2', formula: 'Salty Lychee 5.4 (Giv)', showComplete: true, borderColor: '#d53f8c' },
  { code: '120001681', name: 'Grape Flavor FA-10L @3.31', formula: 'Grape Berry 5.4', showComplete: false, borderColor: '#805ad5' },
  { code: '120001567', name: 'Raspberry Flavor FA-09L @5.51', formula: 'Grape Berry 5.4', showComplete: true, borderColor: '#805ad5' },
  { code: '120001876', name: 'Pink Bomb Strawberry Flavor FA-18L @5.46', formula: 'Pink Bomb Strawberry Lime 5.4', showComplete: true, borderColor: '#fc8181' },
  { code: '120001957', name: 'Sarsi Part 1#4.0', formula: 'SARSI 4.6', showComplete: true, borderColor: '#a0aec0' },
  { code: '120001760', name: 'est Sugar free DryComponents', formula: 'Cola Sugar Free', showComplete: true, borderColor: '#cbd5e0' },
  { code: '120001436', name: 'Orange Emulsion Compound FA-07L @19.635', formula: 'BIB-Orange 5.4', showComplete: true, borderColor: '#f6ad55' },
  { code: '120001437', name: 'Cream Soda Flavor FA-06L @13.23', formula: 'BIB-Cream Soda 5.4', showComplete: true, borderColor: '#48bb78' },
  { code: '120001455', name: 'Est Strawberry Dry Component SF1', formula: 'BIB-Strawberry 5.4', showComplete: true, borderColor: '#fc8181' },
  { code: '120001772', name: 'Est Play Kamikaze Dry Components KMF2', formula: 'BIB-Kamikaze 5.4', showComplete: true, borderColor: '#3182ce' },
  { code: '120002035', name: 'Est Strawberry Dry Components (E) 5.4', formula: 'Strawberry 5.4 (EXPORT)', showComplete: true, borderColor: '#fc8181' },
  { code: '120002003', name: 'Premix Kyoho Sparkling (SS)#HCL2', formula: 'Chakulza Kyoho (CKH)', showComplete: true, borderColor: '#b7791f' },
  { code: '120001959', name: 'Flavor Honey Lemon Sparkling (SS)#HCL2', formula: 'Chakulza Honey Lemon (CLM)', showComplete: true, borderColor: '#ecc94b' },
  { code: '120002221', name: 'Flavor Peach Sparking Can 5t(SS)', formula: 'Flavor Peach Sparking Can 5t(SS)', showComplete: true, borderColor: '#d53f8c' },
  { code: '120001956', name: 'Polysorbate (OT-6L)', formula: '', showComplete: false, borderColor: '#cbd5e0' },
  { code: '120001181', name: 'Regular Part 2A 100 Plus', formula: '', showComplete: false, borderColor: '#cbd5e0' },
  { code: '120001880', name: 'Lime Emulsion Flavor FA-38L @19.43', formula: '', showComplete: false, borderColor: '#38a169' },
  { code: '120001872', name: 'Calamansi flavor FA-40L', formula: '', showComplete: false, borderColor: '#cbd5e0' },
  { code: '120001884', name: 'Coconut flavor FA-41L', formula: '', showComplete: false, borderColor: '#cbd5e0' },
  { code: '120002109', name: '111031 Sweet flavor SW-11L(CIF)', formula: '', showComplete: false, borderColor: '#cbd5e0' },
  { code: '120002117', name: 'Clouding Agent FA-42L', formula: '', showComplete: false, borderColor: '#cbd5e0' },
  { code: '120002103', name: 'Peach Flavour FA-43L', formula: '', showComplete: false, borderColor: '#cbd5e0' },
  { code: '120002108', name: 'Apple Kiwi Flavor FA-45L', formula: 'Electric green', showComplete: true, borderColor: '#48bb78' },
  { code: '120002100', name: 'Cherry blossom FA-44L', formula: 'Pinky winky', showComplete: true, borderColor: '#d53f8c' },
  { code: '120002106', name: 'Mixed berry flavour FA-29L', formula: 'Sigma Blue', showComplete: true, borderColor: '#805ad5' },
  { code: '120002107', name: 'Cantaloupe flavor FA-48L', formula: 'Flashy Yellow', showComplete: true, borderColor: '#ecc94b' },
  { code: '120002031', name: 'BIB Cola sugar free Dry component x 1 unit', formula: 'BIB-Cola sugar free', showComplete: true, borderColor: '#ed8936' },
  { code: '120002030', name: 'BIB Orange Dry component x 1 unit', formula: 'BIB-Orange 5.4', showComplete: true, borderColor: '#f6ad55' },
  { code: '120002032', name: 'BIB Strawberry Dry component x 1 unit', formula: 'BIB-Strawberry 5.4', showComplete: true, borderColor: '#fc8181' },
  { code: '120002033', name: 'BIB Cream soda Dry component x 1 unit', formula: 'BIB-Cream Soda 5.4', showComplete: true, borderColor: '#48bb78' },
  { code: '120002034', name: 'BIB Lemon lime Dry component x 1 unit', formula: 'BIB-Lemon Lime 5.4', showComplete: true, borderColor: '#38a169' }
];

function getActiveStockForCode(code) {
  const latestBalances = new Map();
  const sortedStock = [...state.allStock].sort((a, b) => (b.check_date||'').localeCompare(a.check_date||''));
  
  sortedStock.forEach(s => {
    if (s.code !== code) return;
    const key = `${s.code}__${s.batch}`;
    if (!latestBalances.has(key)) {
      latestBalances.set(key, s.quantity || 0);
    }
  });
  
  let total = 0;
  latestBalances.forEach(qty => {
    if (qty > 0) total += qty;
  });
  return total;
}

function getFormulaBottleneck(formulaName) {
  if (!formulaName) return null;
  
  // Only consider key ingredients that are part of the static report items
  const reportCodes = new Set(staticReportItems.map(item => item.code));
  
  const ingredients = state.allMaster.filter(m => 
    m.product_name === formulaName && 
    m.is_monitored !== false && 
    reportCodes.has(m.code)
  );
  if (ingredients.length === 0) return null;
  
  let minUnits = Infinity;
  ingredients.forEach(ing => {
    const factor = ing.unit_per_pack && ing.unit_per_pack > 0 ? ing.unit_per_pack : 1.0;
    const codeStock = getActiveStockForCode(ing.code);
    const potential = codeStock * factor;
    if (potential < minUnits) {
      minUnits = potential;
    }
  });
  return minUnits === Infinity ? 0 : Math.floor(minUnits);
}

async function openEmailPreviewModal() {
  try {
    const modal = document.getElementById('email-preview-modal');
    if (!modal) {
      console.error('Modal element email-preview-modal not found');
      return;
    }
    
    // Auto-fetch email recipients from database users table
    try {
      const { data, error } = await sb.from('users').select('*').eq('send_email', true);
      if (!error && data) {
        const toList = data.filter(r => r.email_type === 'to').map(r => r.email).join(', ');
        const ccList = data.filter(r => r.email_type === 'cc').map(r => r.email).join(', ');
        
        const toEl = document.getElementById('email-to');
        if (toEl) toEl.value = toList;
        
        const ccEl = document.getElementById('email-cc');
        if (ccEl) ccEl.value = ccList;
      }
    } catch (dbErr) {
      console.warn("Could not load email recipients from users table:", dbErr);
    }
    
    // Format current date and time
    const today = new Date();
    const day = String(today.getDate()).padStart(2, '0');
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const year = today.getFullYear();
    const hours = String(today.getHours()).padStart(2, '0');
    const minutes = String(today.getMinutes()).padStart(2, '0');
    const dateTimeStr = `${day}/${month}/${year} ${hours}:${minutes}`;
    
    const subject = `รายงานหัวเชื้อ / วันหมดอายุ ณ วันที่ ${day}/${month}/${year}`;
    const subjectEl = document.getElementById('email-subject');
    if (subjectEl) subjectEl.value = subject;
    
    // 1. Calculate stock levels and formula bottlenecks
    // DEBUG: log available formula names and stock state
    console.log('[Email Debug] allStock rows:', state.allStock.length, '| allMaster rows:', state.allMaster.length);
    const availableFormulas = [...new Set(state.allMaster.map(m => m.product_name).filter(Boolean))].sort();
    console.log('[Email Debug] product_names in allMaster:', availableFormulas);
    const staticFormulas = [...new Set(staticReportItems.map(i => i.formula).filter(Boolean))];
    const missingFormulas = staticFormulas.filter(f => !availableFormulas.includes(f));
    if (missingFormulas.length) console.warn('[Email Debug] Formulas NOT found in allMaster:', missingFormulas);
    // DEBUG: sample stock codes
    const staticCodes = [...new Set(staticReportItems.map(i => i.code))];
    const stockCodes = [...new Set(state.allStock.map(s => s.code))];
    const missingCodes = staticCodes.filter(c => !stockCodes.includes(c));
    if (missingCodes.length) console.warn('[Email Debug] Codes NOT found in allStock:', missingCodes);

    const itemsWithData = staticReportItems.map(item => {
      const stockQty = getActiveStockForCode(item.code);
      let bottleneck = null;
      let isLimiting = false;
      let potential = null;
      
      if (item.formula) {
        bottleneck = getFormulaBottleneck(item.formula);
        const masterItem = state.allMaster.find(m => m.code === item.code && m.product_name === item.formula);
        const factor = (masterItem && masterItem.unit_per_pack && masterItem.unit_per_pack > 0) ? masterItem.unit_per_pack : 1.0;
        potential = stockQty * factor;
        
        if (bottleneck !== null) {
          if (bottleneck === 0) {
            if (stockQty === 0) {
              isLimiting = true;
            }
          } else {
            if (Math.abs(potential - bottleneck) < 1.0) {
              isLimiting = true;
            }
          }
        }
      }
      
      let sortCompleteValue = -1;
      if (item.formula) {
        sortCompleteValue = (bottleneck !== null && bottleneck > 0) ? bottleneck : 0;
      }
      
      return {
        ...item,
        stockQty,
        bottleneck,
        isLimiting,
        potential,
        sortCompleteValue
      };
    });
    
    // 2. Sort items: Descending complete capability (มาก -> น้อย), then group by formula (limiting ingredient first), then un-grouped at the bottom sorted by stock
    itemsWithData.sort((a, b) => {
      // Compare by complete capability (formula bottleneck)
      // High to low
      if (b.sortCompleteValue !== a.sortCompleteValue) {
        return b.sortCompleteValue - a.sortCompleteValue;
      }
      
      // If same capability, prioritize formula over non-formula
      if (a.formula && !b.formula) return -1;
      if (!a.formula && b.formula) return 1;
      
      // If both have formulas (and they have the same capacity), group by formula name
      if (a.formula && b.formula) {
        if (a.formula !== b.formula) {
          return a.formula.localeCompare(b.formula);
        }
        // Within same formula, put limiting (bottleneck) first
        if (a.isLimiting !== b.isLimiting) {
          return a.isLimiting ? -1 : 1;
        }
      }
      
      // Secondary sort: individual stock level (high to low)
      return b.stockQty - a.stockQty;
    });
    
    // Build HTML Table (Professional styling, modern colors, clean border design)
    let html = ``;
    html += `<div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px;">`;
    html += `  <span style="font-size: 16px; font-weight: 700; color: #0f172a;">รายงานสรุปยอดคงคลังหัวเชื้อและกำลังการผลิต</span>`;
    html += `  <span style="font-size: 12px; color: #475569; background-color: #f1f5f9; padding: 6px 12px; border-radius: 20px; font-weight: 600; border: 1px solid #e2e8f0;">ข้อมูล ณ วันที่ ${dateTimeStr}</span>`;
    html += `</div>`;
    
    html += `<div style="overflow: hidden; border: 1px solid #cbd5e1; border-radius: 8px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);">`;
    html += `<table style="border-collapse: collapse; width: 100%; font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif; font-size: 12.5px; color: #334155; line-height: 1.5; background-color: #ffffff;">`;
    html += `<thead>`;
    html += `  <tr style="background-color: #0f172a; color: #ffffff;">`;
    html += `    <th style="padding: 12px 8px; text-align: center; font-weight: 600; width: 45px; border: 1px solid #cbd5e1; font-size: 11.5px; text-transform: uppercase;">No.</th>`;
    html += `    <th style="padding: 12px 8px; text-align: center; font-weight: 600; width: 90px; border: 1px solid #cbd5e1; font-size: 11.5px; text-transform: uppercase;">Code</th>`;
    html += `    <th style="padding: 12px 12px; text-align: left; font-weight: 600; border: 1px solid #cbd5e1; font-size: 11.5px; text-transform: uppercase;">New Description</th>`;
    html += `    <th style="padding: 12px 12px; text-align: left; font-weight: 600; width: 160px; border: 1px solid #cbd5e1; font-size: 11.5px; text-transform: uppercase;">Formula / Product</th>`;
    html += `    <th style="padding: 12px 12px; text-align: right; font-weight: 600; width: 105px; border: 1px solid #cbd5e1; font-size: 11.5px; text-transform: uppercase;">Flavor/Units</th>`;
    html += `    <th style="padding: 12px 12px; text-align: right; font-weight: 600; width: 150px; border: 1px solid #cbd5e1; font-size: 11.5px; text-transform: uppercase;">Stock/Complete</th>`;
    html += `  </tr>`;
    html += `</thead>`;
    html += `<tbody>`;
    
    itemsWithData.forEach((item, index) => {
      // Determine alternating row bg
      const rowBg = (index % 2 === 0) ? '#ffffff' : '#f8fafc';
      
      // Stock quantity text
      const stockText = (item.stockQty > 0) ? item.stockQty.toLocaleString() : '0';
      const stockColor = (item.stockQty > 0) ? '#0f172a' : '#ef4444';
      const stockFontWeight = (item.stockQty > 0) ? '600' : 'bold';
      
      // Formula badge html
      let formulaHtml = `<span style="color: #94a3b8; font-style: italic;">ทั่วไป</span>`;
      if (item.formula) {
        const badgeColor = item.borderColor || '#64748b';
        const badgeStyle = `background-color: ${badgeColor}15; color: ${badgeColor}; border: 1px solid ${badgeColor}35; padding: 3px 8px; border-radius: 6px; font-weight: 600; font-size: 11px; display: inline-block;`;
        formulaHtml = `<span style="${badgeStyle}">${item.formula}</span>`;
      }
      
      // Stock/Complete badge html
      let completeHtml = `<span style="color: #94a3b8; font-size: 12px;">-</span>`;
      if (item.formula) {
        const bottleneckVal = item.bottleneck !== null ? item.bottleneck : 0;
        
        if (bottleneckVal === 0) {
          if (item.isLimiting) {
            completeHtml = `<span style="background-color: #fee2e2; color: #b91c1c; border: 1px solid #fca5a5; padding: 4px 8px; border-radius: 6px; font-weight: 700; font-size: 11px; display: inline-block; box-shadow: 0 1px 2px rgba(0,0,0,0.02);">⚠️ 0 UNIT (คอขวด)</span>`;
          } else {
            completeHtml = `<span style="background-color: #f1f5f9; color: #64748b; border: 1px solid #cbd5e1; padding: 4px 8px; border-radius: 6px; font-weight: 600; font-size: 11px; display: inline-block;">0 UNIT</span>`;
          }
        } else {
          if (item.isLimiting) {
            completeHtml = `<span style="background-color: #fffbeb; color: #b45309; border: 1px solid #fde68a; padding: 4px 8px; border-radius: 6px; font-weight: 700; font-size: 11px; display: inline-block; box-shadow: 0 1px 2px rgba(0,0,0,0.02);">⚠️ ${bottleneckVal.toLocaleString()} UNIT</span>`;
          } else {
            completeHtml = `<span style="background-color: #f0fdf4; color: #15803d; border: 1px solid #bbf7d0; padding: 4px 8px; border-radius: 6px; font-weight: 700; font-size: 11px; display: inline-block;">${bottleneckVal.toLocaleString()} UNIT</span>`;
          }
        }
      }
      
      // Inline styles for cells
      const baseCellStyle = `border: 1px solid #cbd5e1; padding: 8px 12px; background-color: ${rowBg}; vertical-align: middle;`;
      const codeStyle = `${baseCellStyle} text-align: center; font-family: Consolas, Monaco, monospace; color: #475569; font-size: 12px;`;
      const descStyle = `${baseCellStyle} text-align: left; font-weight: 500; color: #1e293b;`;
      const formulaStyle = `${baseCellStyle} text-align: left;`;
      const stockStyle = `${baseCellStyle} text-align: right; font-family: Consolas, Monaco, monospace; font-weight: ${stockFontWeight}; color: ${stockColor};`;
      const completeStyle = `${baseCellStyle} text-align: right;`;
      
      // Accent color strip on the left-most column (No.)
      const accentStyle = `border: 1px solid #cbd5e1; border-left: 4px solid ${item.borderColor || '#cbd5e0'}; padding: 8px 8px; background-color: ${rowBg}; text-align: center; font-weight: bold; color: #475569; vertical-align: middle;`;
      
      html += `  <tr>`;
      html += `    <td style="${accentStyle}">${index + 1}</td>`;
      html += `    <td style="${codeStyle}">${item.code}</td>`;
      html += `    <td style="${descStyle}">${item.name}</td>`;
      html += `    <td style="${formulaStyle}">${formulaHtml}</td>`;
      html += `    <td style="${stockStyle}">${stockText}</td>`;
      html += `    <td style="${completeStyle}">${completeHtml}</td>`;
      html += `  </tr>`;
    });
    
    html += `</tbody>`;
    html += `</table>`;
    html += `</div>`;
    
    const previewEl = document.getElementById('email-body-preview');
    if (previewEl) previewEl.innerHTML = html;

    modal.classList.add('open');
  } catch (err) {
    console.error('Error generating email preview:', err);
    showToast('เกิดข้อผิดพลาดในการสร้างพรีวิว: ' + err.message, 'error');
  }
}

function closeEmailPreviewModal() {
  const modal = document.getElementById('email-preview-modal');
  if (modal) modal.classList.remove('open');
}

async function openOutlookWeb() {
  try {
    const to = document.getElementById('email-to').value;
    const subject = document.getElementById('email-subject').value;
    
    // 1. Copy the HTML table to clipboard first (so user can still use rich paste if they want to)
    const tableEl = document.querySelector('#email-body-preview table');
    if (tableEl) {
      try {
        const type = "text/html";
        const blob = new Blob([tableEl.outerHTML], { type });
        const data = [new ClipboardItem({ [type]: blob })];
        await navigator.clipboard.write(data);
      } catch (clipErr) {
        console.error('Auto-copy table failed:', clipErr);
        try {
          await navigator.clipboard.writeText(tableEl.innerText);
        } catch(fErr) {}
      }
    }

    // 2. Generate ASCII text representation of the table for pre-populating Outlook Web body
    const today = new Date();
    const day = String(today.getDate()).padStart(2, '0');
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const year = today.getFullYear();
    const hours = String(today.getHours()).padStart(2, '0');
    const minutes = String(today.getMinutes()).padStart(2, '0');
    const dateTimeStr = `${day}/${month}/${year} ${hours}:${minutes}`;

    let body = `เรียน ทีมงานผลิตและคลังสินค้า\n\n`;
    body += `สรุปกำลังการผลิต Complete Set ณ วันที่ ${dateTimeStr}\n\n`;
    body += `Item | Code      | New Description                           | Formula                      | Flavor/Units | Stock/Complete\n`;
    body += `----------------------------------------------------------------------------------------------------------------------\n`;

    const tableRows = document.querySelectorAll('#email-body-preview table tbody tr');
    tableRows.forEach(row => {
      const cells = row.querySelectorAll('td');
      if (cells.length >= 6) {
        const item = cells[0].innerText.trim().padEnd(4);
        const code = cells[1].innerText.trim().padEnd(9);
        const desc = cells[2].innerText.trim().padEnd(41);
        const formula = cells[3].innerText.trim().padEnd(28);
        const qty = cells[4].innerText.trim().padStart(12);
        const complete = cells[5].innerText.trim().padStart(20);
        body += `${item} | ${code} | ${desc} | ${formula} | ${qty} | ${complete}\n`;
      }
    });

    body += `\n==================================================\n`;
    body += `ข้อมูลนี้จัดทำขึ้นโดยระบบติดตามวันหมดอายุหัวเชื้อ StockFlow\n`;
    body += `(หมายเหตุ: ระบบได้คัดลอกตารางสีสันลง Clipboard แล้ว สามารถกด Ctrl+V เพื่อวางทับเป็นแบบสวยงามได้ตามต้องการ)`;
    
    // 3. Open Outlook Web compose link with the pre-populated body
    const outlookUrl = `https://outlook.office.com/mail/deeplink/compose?to=${encodeURIComponent(to)}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
    window.open(outlookUrl, '_blank');
    
    showToast('กำลังเปิด Outlook Web พร้อมรายละเอียดสรุป (สามารถกดส่งได้ทันที หรือกด Ctrl+V เพื่อวางตารางสีสัน)', 'success');
  } catch (err) {
    showToast('เกิดข้อผิดพลาดในการเปิด Outlook Web: ' + err.message, 'error');
  }
}

async function copyTableToClipboard() {
  const tableEl = document.querySelector('#email-body-preview table');
  if (!tableEl) return;
  try {
    const type = "text/html";
    const blob = new Blob([tableEl.outerHTML], { type });
    const data = [new ClipboardItem({ [type]: blob })];
    await navigator.clipboard.write(data);
    showToast('คัดลอกตารางไปยัง Clipboard สำเร็จ! สามารถนำไปวางใน Outlook ได้ทันที', 'success');
  } catch (err) {
    console.error('HTML clipboard copy failed, falling back to text:', err);
    try {
      // Fallback: Copy as plaintext
      await navigator.clipboard.writeText(tableEl.innerText);
      showToast('คัดลอกข้อความตารางสำเร็จ (ไม่ติดรูปแบบ)', 'success');
    } catch (fallbackErr) {
      showToast('ไม่สามารถคัดลอกได้: ' + fallbackErr.message, 'error');
    }
  }
}

// ---- EXCEL UPLOAD AND DATA PROCESSING ----
function triggerExcelUpload() {
  const el = document.getElementById('excel-upload-input');
  if (el) el.click();
}

function parseExcelDate(val) {
  if (!val) return null;
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null;
    const year = val.getFullYear();
    const month = String(val.getMonth() + 1).padStart(2, '0');
    const day = String(val.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }
  let s = String(val).trim();
  if (!s || s.toLowerCase() === 'nan' || s.toLowerCase() === 'none' || s === '0') return null;
  
  // Handle DD/MM/YYYY or DD.MM.YYYY or DD-MM-YYYY
  const parts = s.split(/[\/\-\.]/);
  if (parts.length === 3) {
    let day = parseInt(parts[0], 10);
    let month = parseInt(parts[1], 10);
    let year = parseInt(parts[2], 10);
    if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
      // Handle 2 digit year
      if (year < 100) year += 2000;
      // Handle Buddhist Era (BE) e.g., 2568 -> 2025
      if (year > 2400) year -= 543;
      
      const d = new Date(year, month - 1, day);
      if (!isNaN(d.getTime())) {
        return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      }
    }
  }
  
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }
  return null;
}

async function handleExcelUpload(event) {
  const file = event.target.files[0];
  if (!file) return;
  
  const message = "⚠️ คำเตือน!\nการอัปโหลดไฟล์ Excel ใหม่จะทำการลบข้อมูล Stock ปัจจุบันทั้งหมดในฐานข้อมูล และแทนที่ด้วยข้อมูลจากไฟล์นี้เท่านั้น\n\nต้องการเริ่มนำเข้าข้อมูลหรือไม่?";
  if (!confirm(message)) {
    event.target.value = '';
    return;
  }
  
  const modal = document.getElementById('upload-progress-modal');
  const pBar = document.getElementById('upload-progress-bar');
  const pTitle = document.getElementById('upload-progress-title');
  const pText = document.getElementById('upload-progress-text');
  
  if (modal) modal.classList.add('open');
  if (pBar) pBar.style.width = '5%';
  if (pTitle) pTitle.textContent = "กำลังเตรียมนำเข้า...";
  if (pText) pText.textContent = "กำลังประมวลผลไฟล์ Excel...";
  
  const reader = new FileReader();
  reader.onload = async function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });
      
      const skipSheets = ["Code", "Code (2)", "Aging"];
      const allRecords = [];
      
      for (const sheetName of workbook.SheetNames) {
        if (skipSheets.includes(sheetName)) continue;
        
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) continue;
        
        // Convert to 2D array
        const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });
        
        // Find header row containing 'Code'
        let headerRowIdx = -1;
        for (let i = 0; i < sheetData.length; i++) {
          const row = sheetData[i];
          if (row && row.some(v => v !== null && v !== undefined && String(v).trim() === 'Code')) {
            headerRowIdx = i;
            break;
          }
        }
        
        if (headerRowIdx === -1) {
          console.log(`Skipped sheet ${sheetName}: no 'Code' header row found.`);
          continue;
        }
        
        // Find check date (first date cell above header row)
        let checkDate = null;
        for (let i = 0; i < headerRowIdx; i++) {
          const row = sheetData[i];
          if (!row) continue;
          for (const val of row) {
            if (val instanceof Date && !isNaN(val.getTime())) {
              checkDate = parseExcelDate(val);
              break;
            }
          }
          if (checkDate) break;
        }
        
        const headerRow = sheetData[headerRowIdx];
        
        // Map columns
        const colMap = {};
        headerRow.forEach((h, i) => {
          if (h === null || h === undefined) return;
          const name = String(h).trim();
          if (name === "Code") colMap["code"] = i;
          else if (name === "Description") colMap["description"] = i;
          else if (name.includes("New Description")) {
            if (!colMap["new_description"]) colMap["new_description"] = i;
          }
          else if (name.includes("อายุสินค้า")) colMap["shelf_life"] = i;
          else if (name.includes("Batch") || name.toLowerCase().includes("batch")) colMap["batch"] = i;
          else if (name.includes("รับสินค้า")) colMap["received_date"] = i;
          else if (name.includes("ผลิต") && !name.includes("รับ")) colMap["production_date"] = i;
          else if (name.includes("หมดอายุ") && !name.includes("ต่ออายุ") && !name.includes("ใกล้")) colMap["expiry_date"] = i;
          else if (name.includes("ต่ออายุ1")) colMap["extended_expiry_1"] = i;
          else if (name.includes("ต่ออายุ2")) colMap["extended_expiry_2"] = i;
          else if (name.includes("จำนวน")) colMap["quantity"] = i;
        });
        
        let currentCode = null;
        let currentDesc = null;
        let currentNewDesc = null;
        let currentShelf = null;
        
        // Read rows below header
        for (let idx = headerRowIdx + 1; idx < sheetData.length; idx++) {
          const rowVals = sheetData[idx];
          if (!rowVals || rowVals.length === 0) continue;
          
          // Update carried values
          if (colMap["code"] !== undefined) {
            let v = rowVals[colMap["code"]];
            if (v !== null && v !== undefined) {
              v = String(v).trim();
              if (v) {
                if (v.includes('.')) {
                  v = v.split('.')[0];
                }
                currentCode = v;
                currentDesc = null;
                currentNewDesc = null;
              }
            }
          }
          
          if (colMap["description"] !== undefined) {
            const v = rowVals[colMap["description"]];
            if (v !== null && v !== undefined && String(v).trim()) {
              currentDesc = String(v).trim();
            }
          }
          
          if (colMap["new_description"] !== undefined) {
            const v = rowVals[colMap["new_description"]];
            if (v !== null && v !== undefined && String(v).trim()) {
              currentNewDesc = String(v).trim();
            }
          }
          
          if (colMap["shelf_life"] !== undefined) {
            const v = rowVals[colMap["shelf_life"]];
            if (v !== null && v !== undefined && String(v).trim()) {
              currentShelf = String(v).trim();
            }
          }
          
          if (!currentCode) continue;
          
          let batchVal = null;
          if (colMap["batch"] !== undefined) {
            const v = rowVals[colMap["batch"]];
            if (v !== null && v !== undefined) {
              batchVal = String(v).trim();
            }
          }
          if (!batchVal) batchVal = 'no-batch';
          
          // Clean batch trailing .0
          if (batchVal.includes('.') && batchVal.split('.')[1] === '0') {
            batchVal = batchVal.split('.')[0];
          }
          
          let qty = null;
          if (colMap["quantity"] !== undefined) {
            const v = rowVals[colMap["quantity"]];
            if (v !== null && v !== undefined) {
              qty = parseFloat(v);
            }
          }
          
          // Skip quantity <= 0
          if (qty === null || isNaN(qty) || qty <= 0) continue;
          
          const record = {
            product_group:      sheetName.trim(),
            code:               currentCode,
            description:        currentDesc,
            new_description:    currentNewDesc,
            shelf_life:         currentShelf,
            batch:              batchVal,
            received_date:      colMap["received_date"] !== undefined ? parseExcelDate(rowVals[colMap["received_date"]]) : null,
            production_date:    colMap["production_date"] !== undefined ? parseExcelDate(rowVals[colMap["production_date"]]) : null,
            expiry_date:        colMap["expiry_date"] !== undefined ? parseExcelDate(rowVals[colMap["expiry_date"]]) : null,
            extended_expiry_1:  colMap["extended_expiry_1"] !== undefined ? parseExcelDate(rowVals[colMap["extended_expiry_1"]]) : null,
            extended_expiry_2:  colMap["extended_expiry_2"] !== undefined ? parseExcelDate(rowVals[colMap["extended_expiry_2"]]) : null,
            quantity:           qty,
            check_date:         checkDate || new Date().toISOString().split('T')[0],
          };
          
          allRecords.push(record);
        }
      }
      
      if (allRecords.length === 0) {
        alert("❌ ไม่พบข้อมูลวัตถุดิบหัวเชื้อที่สามารถนำเข้าได้จากไฟล์นี้ กรุณาตรวจสอบชีทและคอลัมน์");
        if (modal) modal.classList.remove('open');
        event.target.value = '';
        return;
      }
      
      // ⚠️ ยืนยันการลบข้อมูลทั้งหมดก่อน import — ป้องกันข้อมูลหายโดยไม่ตั้งใจ
      if (modal) modal.classList.remove('open');
      const confirmDelete = confirm(
        `⚠️ คำเตือน: การ Import จะลบข้อมูล Stock ทั้งหมดในฐานข้อมูล แล้วแทนที่ด้วยข้อมูลจากไฟล์ Excel นี้\n\n` +
        `พบข้อมูลในไฟล์: ${allRecords.length} แถว\n\n` +
        `การดำเนินการนี้ไม่สามารถย้อนกลับได้ ต้องการดำเนินการต่อหรือไม่?`
      );
      if (!confirmDelete) {
        event.target.value = '';
        return;
      }
      if (modal) modal.classList.add('open');

      if (pTitle) pTitle.textContent = "กำลังล้างข้อมูลเดิม...";
      if (pText) pText.textContent = `พบข้อมูลทั้งหมด ${allRecords.length} แถว กำลังเคลียร์ข้อมูลเดิมในฐานข้อมูล...`;

      // Delete old stock entries
      const { error: deleteError } = await sb.from('stock_entries').delete().neq('id', '00000000-0000-0000-0000-000000000000');
      if (deleteError) {
        throw new Error("ลบข้อมูลเดิมล้มเหลว: " + deleteError.message);
      }
      
      // Insert in batches of 100
      const batchSize = 100;
      const totalBatches = Math.ceil(allRecords.length / batchSize);
      
      for (let b = 0; b < totalBatches; b++) {
        const start = b * batchSize;
        const end = start + batchSize;
        const batchData = allRecords.slice(start, end);
        
        if (pTitle) pTitle.textContent = `กำลังบันทึกข้อมูล... [${b+1}/${totalBatches}]`;
        if (pText) pText.textContent = `กำลังเขียนข้อมูลแถวที่ ${start + 1} ถึง ${Math.min(end, allRecords.length)} จาก ${allRecords.length} แถว`;
        if (pBar) pBar.style.width = `${Math.round(((b + 1) / totalBatches) * 100)}%`;
        
        const { error: insertError } = await sb.from('stock_entries').insert(batchData);
        if (insertError) {
          throw new Error(`นำเข้าชุดที่ ${b+1} ล้มเหลว: ` + insertError.message);
        }
      }
      
      if (modal) modal.classList.remove('open');
      alert(`🎉 นำเข้าข้อมูลคลังสินค้าสำเร็จทั้งหมด ${allRecords.length} แถว!`);
      event.target.value = '';
      
      // Reload and re-render stock page
      await refreshAll();
      
    } catch (err) {
      console.error(err);
      if (modal) modal.classList.remove('open');
      alert("❌ เกิดข้อผิดพลาดในการประมวลผลไฟล์:\n" + err.message);
      event.target.value = '';
    }
  };
  reader.readAsArrayBuffer(file);
}

// ---- USER MANAGEMENT PAGE ----
async function renderUsersPage() {
  const tbody = document.getElementById('users-tbody');
  if (!tbody) return;
  
  tbody.innerHTML = `<tr><td colspan="6" style="text-align: center; padding: 24px; color: var(--muted);"><span class="spinner" style="display:inline-block; margin-right:8px;"></span>กำลังโหลดรายชื่อผู้ใช้งาน...</td></tr>`;
  
  try {
    const { data, error } = await sb.from('users').select('*').order('created_at', { ascending: false });
    if (error) throw error;
    
    state.allUsers = data || [];
    
    if (state.allUsers.length === 0) {
      tbody.innerHTML = `<tr><td colspan="6" style="text-align: center; padding: 24px; color: var(--muted);">ไม่พบข้อมูลผู้ใช้งานในระบบ</td></tr>`;
      return;
    }
    
    tbody.innerHTML = state.allUsers.map(u => {
      const isChecked = u.send_email ? 'checked' : '';
      const typeOptions = `
        <select onchange="updateUserEmailType('${u.id}', this.value)" style="padding: 4px 8px; font-size:12px; border:1px solid var(--border); border-radius:4px; background:#fff;">
          <option value="to" ${u.email_type === 'to' ? 'selected' : ''}>To (ผู้รับหลัก)</option>
          <option value="cc" ${u.email_type === 'cc' ? 'selected' : ''}>CC (สำเนา)</option>
        </select>
      `;
      
      const roleMap = {
        'admin': 'Admin (ผู้ดูแลระบบ)',
        'operator': 'Operator (ผู้บันทึก)',
        'viewer': 'Viewer (ผู้ดูข้อมูล)'
      };
      const roleLabel = roleMap[u.role] || u.role || '—';
      
      return `
        <tr>
          <td><strong>${u.name || '—'}</strong></td>
          <td class="code-text">${u.email}</td>
          <td>${roleLabel}</td>
          <td style="text-align: center;">
            <input type="checkbox" onchange="toggleUserSendEmail('${u.id}', this.checked)" ${isChecked} style="width: 16px; height: 16px; cursor: pointer; accent-color: var(--primary);">
          </td>
          <td>${typeOptions}</td>
          <td style="text-align: center;">
            <button class="btn btn-ghost btn-sm" onclick="openEditUser('${u.id}')" style="color: var(--accent); padding: 4px; height: auto; margin-right: 6px;" title="แก้ไข">
              <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" width="14" height="14" style="stroke-width:2">
                <path stroke-linecap="round" stroke-linejoin="round" d="M16.862 4.487l1.687-1.688a1.875 1.875 0 112.652 2.652L6.832 19.82a4.5 4.5 0 01-1.897 1.13l-2.685.8.8-2.685a4.5 4.5 0 011.13-1.897L16.863 4.487zm0 0L19.5 7.125" />
              </svg>
            </button>
            <button class="btn btn-ghost btn-sm" onclick="deleteUser('${u.id}')" style="color: var(--red); padding: 4px; height: auto;" title="ลบ">
              <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" width="14" height="14" style="stroke-width:2">
                <path stroke-linecap="round" stroke-linejoin="round" d="M14.74 9l-.346 9m-4.788 0L9.26 9m9.968-3.21c.342.052.682.107 1.022.166m-1.022-.165L18.16 19.673a2.25 2.25 0 01-2.244 2.077H8.084a2.25 2.25 0 01-2.244-2.077L4.772 5.79m14.456 0a48.108 48.108 0 00-3.478-.397m-12 .562c.34-.059.68-.114 1.022-.165m0 0a48.11 48.11 0 013.478-.397m7.5 0v-.916c0-1.18-.91-2.164-2.09-2.201a51.964 51.964 0 00-3.32 0c-1.18.037-2.09 1.022-2.09 2.201v.916m7.5 0a48.667 48.667 0 00-7.5 0" />
              </svg>
            </button>
          </td>
        </tr>
      `;
    }).join('');
  } catch (err) {
    console.error(err);
    tbody.innerHTML = `<tr><td colspan="6" style="text-align: center; padding: 24px; color: var(--red);">เกิดข้อผิดพลาดในการโหลดรายชื่อ: ${err.message}</td></tr>`;
  }
}

async function addUser(event) {
  event.preventDefault();
  
  const nameEl = document.getElementById('u-name');
  const emailEl = document.getElementById('u-email');
  const roleEl = document.getElementById('u-role');
  const passwordEl = document.getElementById('u-password');
  const emailTypeEl = document.getElementById('u-email-type');
  const sendEmailEl = document.getElementById('u-send-email');
  
  if (!nameEl || !emailEl || !roleEl || !emailTypeEl || !sendEmailEl) return;
  
  const name = nameEl.value.trim();
  const email = emailEl.value.trim();
  const role = roleEl.value;
  const password = passwordEl ? passwordEl.value : '';
  const email_type = emailTypeEl.value;
  const send_email = sendEmailEl.checked;
  
  try {
    const { error } = await sb.from('users').insert({ name, email, role, password, email_type, send_email });
    if (error) throw error;
    
    // Clear form
    nameEl.value = '';
    emailEl.value = '';
    if (passwordEl) passwordEl.value = '';
    roleEl.value = 'operator';
    emailTypeEl.value = 'to';
    sendEmailEl.checked = true;
    
    showToast('เพิ่มข้อมูลผู้ใช้งานเรียบร้อยแล้ว', 'success');
    await renderUsersPage();
  } catch (err) {
    console.error(err);
    alert('❌ บันทึกผู้ใช้ล้มเหลว:\n' + err.message);
  }
}

async function deleteUser(id) {
  if (!confirm("⚠️ ยืนยันการลบผู้ใช้งานรายนี้ออกจากระบบ?")) return;
  
  try {
    const { error } = await sb.from('users').delete().eq('id', id);
    if (error) throw error;
    
    showToast('ลบข้อมูลผู้ใช้งานสำเร็จ', 'success');
    await renderUsersPage();
  } catch (err) {
    console.error(err);
    alert('❌ ลบผู้ใช้ล้มเหลว:\n' + err.message);
  }
}

async function toggleUserSendEmail(id, status) {
  try {
    const { error } = await sb.from('users').update({ send_email: status }).eq('id', id);
    if (error) throw error;
    showToast('อัปเดตสิทธิ์การรับรายงานทางอีเมลเรียบร้อย', 'success');
  } catch (err) {
    console.error(err);
    showToast('เกิดข้อผิดพลาดในการอัปเดตสถานะ: ' + err.message, 'error');
  }
}

async function updateUserEmailType(id, type) {
  try {
    const { error } = await sb.from('users').update({ email_type: type }).eq('id', id);
    if (error) throw error;
    showToast('อัปเดตประเภทการรับรายงานเรียบร้อย', 'success');
  } catch (err) {
    console.error(err);
    showToast('เกิดข้อผิดพลาดในการอัปเดตประเภท: ' + err.message, 'error');
  }
}

// ---- EDIT USER FUNCTIONS ----
function openEditUser(id) {
  const u = (state.allUsers || []).find(user => user.id === id);
  if (!u) return;
  
  const idEl = document.getElementById('edit-u-id');
  const nameEl = document.getElementById('edit-u-name');
  const emailEl = document.getElementById('edit-u-email');
  const roleEl = document.getElementById('edit-u-role');
  const passwordEl = document.getElementById('edit-u-password');
  const emailTypeEl = document.getElementById('edit-u-email-type');
  const sendEmailEl = document.getElementById('edit-u-send-email');
  const modalEl = document.getElementById('edit-user-modal');
  
  if (idEl) idEl.value = u.id;
  if (nameEl) nameEl.value = u.name || '';
  if (emailEl) emailEl.value = u.email || '';
  if (roleEl) roleEl.value = u.role || 'operator';
  if (passwordEl) passwordEl.value = ''; // Clear password field on open
  if (emailTypeEl) emailTypeEl.value = u.email_type || 'to';
  if (sendEmailEl) sendEmailEl.checked = !!u.send_email;
  
  if (modalEl) modalEl.classList.add('open');
}

function closeUserModal() {
  const modalEl = document.getElementById('edit-user-modal');
  if (modalEl) modalEl.classList.remove('open');
}

async function saveUserEdit() {
  const id = document.getElementById('edit-u-id')?.value;
  const name = document.getElementById('edit-u-name')?.value.trim();
  const email = document.getElementById('edit-u-email')?.value.trim();
  const role = document.getElementById('edit-u-role')?.value;
  const password = document.getElementById('edit-u-password')?.value;
  const email_type = document.getElementById('edit-u-email-type')?.value;
  const send_email = document.getElementById('edit-u-send-email')?.checked;
  
  if (!name || !email) {
    showToast('กรุณากรอกข้อมูลให้ครบถ้วน', 'error');
    return;
  }
  
  const updateData = { name, email, role, email_type, send_email };
  if (password && password.trim() !== '') {
    updateData.password = password;
  }
  
  try {
    const { error } = await sb.from('users').update(updateData).eq('id', id);
    if (error) throw error;
    
    closeUserModal();
    showToast('อัปเดตข้อมูลผู้ใช้งานเรียบร้อยแล้ว', 'success');
    await renderUsersPage();
    
    // If the edited user is the currently logged-in user, update localStorage and reload!
    if (email === localStorage.getItem('currentUserEmail')) {
      localStorage.setItem('currentUserName', name);
      localStorage.setItem('currentUserRole', role);
      location.reload();
    }
  } catch (err) {
    console.error(err);
    alert('❌ อัปเดตข้อมูลผู้ใช้งานล้มเหลว:\n' + err.message);
  }
}

// ---- STOCK DEDUCTION HISTORY ----
async function loadHistory() {
  const tbody = document.getElementById('hs-tbody');
  if (tbody) {
    tbody.innerHTML = '<tr><td colspan="7" class="loading"><div class="spinner"></div> กำลังโหลดข้อมูล...</td></tr>';
  }
  try {
    const { data, error } = await sb
      .from('stock_entries')
      .select('*')
      .lt('quantity', 0)
      .order('check_date', { ascending: false })
      .order('created_at', { ascending: false });
    
    if (error) throw error;
    state.allHistory = data || [];
    state.historyFiltered = [...state.allHistory];
    
    // Populate date filter dropdown
    const dateFilter = document.getElementById('hs-date-filter');
    if (dateFilter) {
      const dates = [...new Set(state.allHistory.map(h => h.check_date).filter(Boolean))].sort().reverse();
      const firstOption = dateFilter.options[0];
      dateFilter.innerHTML = '';
      dateFilter.appendChild(firstOption);
      dates.forEach(d => {
        const opt = document.createElement('option');
        opt.value = d;
        opt.textContent = fmtDate(new Date(d));
        dateFilter.appendChild(opt);
      });
    }
    
    filterHistory();
  } catch (err) {
    console.error(err);
    if (tbody) {
      tbody.innerHTML = `<tr><td colspan="7" style="text-align:center; color:var(--red);">เกิดข้อผิดพลาดในการโหลดข้อมูล: ${err.message}</td></tr>`;
    }
  }
}

function filterHistory() {
  const query = (document.getElementById('hs-search')?.value || '').toLowerCase().trim();
  const dateVal = document.getElementById('hs-date-filter')?.value || '';
  const factoryVal = document.getElementById('hs-factory-filter')?.value || '';
  
  state.historyFiltered = state.allHistory.filter(h => {
    if (dateVal && h.check_date !== dateVal) return false;
    
    if (factoryVal) {
      const notes = h.notes || '';
      if (factoryVal === 'PT') {
        if (!notes.includes('(ทั่วไป)') && !notes.includes('(PT)')) return false;
      } else {
        if (!notes.includes(`(${factoryVal})`)) return false;
      }
    }
    
    if (query) {
      const code = (h.code || '').toLowerCase();
      const desc = (h.description || '').toLowerCase();
      const rmName = (h.rm_full_name || '').toLowerCase();
      const grp = (h.product_group || '').toLowerCase();
      const batch = (h.batch || '').toLowerCase();
      const notes = (h.notes || '').toLowerCase();
      return code.includes(query) || desc.includes(query) || rmName.includes(query) || grp.includes(query) || batch.includes(query) || notes.includes(query);
    }
    return true;
  });
  
  state.historyPage = 1;
  renderHistoryTable();
}

function sortHistory(field) {
  if (state.historySortField === field) {
    state.historySortAsc = !state.historySortAsc;
  } else {
    state.historySortField = field;
    state.historySortAsc = true;
  }
  renderHistoryTable();
}

function renderHistoryTable() {
  // Render headers
  const thead = document.querySelector('#page-history table thead');
  if (thead) {
    const cols = [
      { label: 'วันที่ตัดจ่าย', field: 'check_date' },
      { label: 'SAP Code', field: 'code' },
      { label: 'ชื่อวัตถุดิบ (ในสูตร)', field: 'rm_full_name' },
      { label: 'สูตรผลิต/กลุ่ม', field: 'product_group' },
      { label: 'Batch', field: 'batch' },
      { label: 'จำนวนที่เบิก', field: 'quantity' },
      { label: 'รายละเอียดการเบิก & ผู้บันทึก', field: 'notes' }
    ];
    
    thead.innerHTML = `<tr>${cols.map(c => {
      const isSorted = state.historySortField === c.field;
      const icon = isSorted ? (state.historySortAsc ? ' ▲' : ' ▼') : ' ↕';
      const savedWidth = state.historyColWidths[c.field];
      const widthStyle = savedWidth ? `width: ${savedWidth}px;` : '';
      return `<th data-field="${c.field}" onclick="sortHistory('${c.field}')" class="sortable-th ${isSorted ? 'sorted-active' : ''}" style="cursor:pointer; user-select:none; position:relative; text-align:center; ${widthStyle}">
        <div class="th-content" style="display:inline-flex; align-items:center; gap:4px; justify-content:center; width:100%;">
          ${c.label} <span class="sort-icon" style="font-size:10px; color:${isSorted ? 'var(--primary)' : 'var(--muted)'};">${icon}</span>
        </div>
      </th>`;
    }).join('')}</tr>`;
  }
  
  // Sort data
  const data = [...state.historyFiltered];
  if (state.historySortField) {
    const f = state.historySortField;
    const asc = state.historySortAsc ? 1 : -1;
    data.sort((a, b) => {
      let valA = a[f];
      let valB = b[f];
      if (f === 'quantity') {
        valA = parseFloat(a[f]) || 0;
        valB = parseFloat(b[f]) || 0;
      } else {
        valA = (valA || '').toString().toLowerCase();
        valB = (valB || '').toString().toLowerCase();
      }
      if (valA < valB) return -1 * asc;
      if (valA > valB) return 1 * asc;
      return 0;
    });
  }
  
  // Paginate
  const total = data.length;
  const start = (state.historyPage - 1) * state.perPage;
  const pageData = data.slice(start, start + state.perPage);
  
  // Render body
  const tbody = document.getElementById('hs-tbody');
  if (tbody) {
    if (pageData.length === 0) {
      tbody.innerHTML = `<tr><td colspan="7" style="padding: 40px; text-align: center; color: var(--muted);">ไม่มีประวัติการตัดสต็อก</td></tr>`;
    } else {
      tbody.innerHTML = pageData.map(h => {
        const qtyVal = parseFloat(h.quantity) || 0;
        const qtyFormatted = qtyVal.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 });
        const dateStr = h.check_date ? fmtDate(new Date(h.check_date)) : '—';
        
        // Lookup official raw material name from cached Master database
        const masterItem = state.allMaster.find(m => m.code === h.code);
        const rmName = masterItem ? masterItem.name : (h.new_description || h.description || '—');
        
        return `
          <tr>
            <td style="text-align:center;">${dateStr}</td>
            <td style="text-align:center; font-family: var(--mono);">${h.code || '—'}</td>
            <td style="text-align:left;">${rmName}</td>
            <td style="text-align:left;">${h.product_group || '—'}</td>
            <td style="text-align:center; font-family: var(--mono);">${h.batch || '—'}</td>
            <td style="text-align:right; font-family: var(--mono); color: var(--red); font-weight: 600;">${qtyFormatted}</td>
            <td style="text-align:left;">${h.notes || '—'}</td>
          </tr>
        `;
      }).join('');
    }
  }
  
  // Render pagination
  renderPagination('hs', total, state.historyPage, (p) => {
    state.historyPage = p;
    renderHistoryTable();
  });
  
  // Resizable columns
  const table = document.querySelector('#page-history table');
  if (table) {
    makeTableResizable(table, state.historyColWidths);
  }
}


