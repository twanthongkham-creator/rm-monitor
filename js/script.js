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
    .select('id,code,check_date,quantity,expiry_date,extended_expiry_1,extended_expiry_2,batch,received_date,rm_full_name,new_description,description,product_name')
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
  // --- SessionStorage cache (5 นาที) เพื่อลด egress ---
  const MASTER_CACHE_KEY = 'supabase_master_v1';
  const MASTER_CACHE_TTL = 5 * 60 * 1000;
  let fromCache = false;
  try {
    const cached = sessionStorage.getItem(MASTER_CACHE_KEY);
    if (cached) {
      const parsed = JSON.parse(cached);
      if (Date.now() - parsed.time < MASTER_CACHE_TTL) {
        state.allMaster = parsed.data;
        fromCache = true;
      }
    }
  } catch(e) {}

  if (!fromCache) {
    const { data, error } = await sb.from('raw_materials').select('id,code,name,product_name,rm_code,recipe_step,is_monitored').order('code');
    if (error) throw error;
    state.allMaster = data || [];
    try {
      sessionStorage.setItem(MASTER_CACHE_KEY, JSON.stringify({ data: state.allMaster, time: Date.now() }));
    } catch(e) {}
  }
  
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
  
  let query = sb.from('stock_with_expiry').select('id,code,check_date,quantity,expiry_date,extended_expiry_1,extended_expiry_2,batch,received_date,rm_full_name,new_description,description,product_name').gt('quantity', 0);
  
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
    : '<tr><td colspan="7" class="empty">ย