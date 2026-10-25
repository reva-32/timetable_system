const _nativeFetch=window.fetch.bind(window);
let _csrfToken=null;
async function _getCsrfToken(){
  if(_csrfToken) return _csrfToken;
  const r=await _nativeFetch('/csrf-token',{credentials:'include'});
  if(!r.ok) throw new Error('Unable to initialize security token');
  const d=await r.json(); _csrfToken=d.csrf_token; return _csrfToken;
}
window.fetch=async (input,init={})=>{
  const method=String(init.method||'GET').toUpperCase();
  const opts={...init,credentials:'include',headers:new Headers(init.headers||{})};
  if(['POST','PUT','PATCH','DELETE'].includes(method) && !String(input).includes('/csrf-token')){
    opts.headers.set('X-CSRF-Token',await _getCsrfToken());
  }
  const response=await _nativeFetch(input,opts);
  if(String(input).includes('/login') && response.ok) _csrfToken=null;
  return response;
};
    function parseCustomDate(dateStr) {
      if (!dateStr) return new Date(NaN);
      if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
        return new Date(dateStr);
      }
      const match = String(dateStr).match(/^(\d{1,2})[-/\s]([A-Za-z]{3})[-/\s](\d{4})$/);
      if (match) {
        const day = parseInt(match[1], 10);
        const monthStr = match[2].toLowerCase();
        const year = parseInt(match[3], 10);
        const months = {
          jan:0, feb:1, mar:2, apr:3, may:4, jun:5,
          jul:6, aug:7, sep:8, oct:9, nov:10, dec:11
        };
        const month = months[monthStr.substring(0, 3)];
        if (month !== undefined) {
          return new Date(year, month, day);
        }
      }
      return new Date(dateStr);
    }

    const uzone = document.getElementById('uzone'), fi = document.getElementById('fileInput'),
      fpill = document.getElementById('fpill'), fname = document.getElementById('fname'),
      genBtn = document.getElementById('genBtn');

    uzone.addEventListener('click', () => fi.click());
    uzone.addEventListener('dragover', e => { e.preventDefault(); uzone.classList.add('drag') });
    uzone.addEventListener('dragleave', () => uzone.classList.remove('drag'));
    uzone.addEventListener('drop', e => {
      e.preventDefault(); uzone.classList.remove('drag');
      const f = e.dataTransfer.files[0];
      if (f && f.name.endsWith('.xlsx')) setFile(f); else alert('Please upload a .xlsx file');
    });
    fi.addEventListener('change', () => fi.files[0] && setFile(fi.files[0]));

    function setFile(f) {
      fname.textContent = f.name;
      fpill.style.display = 'flex';
      uzone.style.display = 'none';
      genBtn.disabled = false;
    }
    function clearFile() {
      fi.value = '';
      fpill.style.display = 'none';
      uzone.style.display = 'flex';
      genBtn.disabled = true;
    }

    const stages = [
      { t: 'Building conflict graph…', p: 12 },
      { t: 'Running DSATUR coloring…', p: 30 },
      { t: 'Allocating rooms greedily…', p: 52 },
      { t: 'Running bipartite matching…', p: 74 },
      { t: 'Validating all constraints…', p: 90 },
      { t: 'Compiling final schedule…', p: 100 },
    ];
    function startLoad() {
      const el = document.getElementById('lstage'), bar = document.getElementById('lbar');
      let i = 0;
      const iv = setInterval(() => {
        if (i >= stages.length) { clearInterval(iv); return }
        el.textContent = stages[i].t;
        bar.style.width = stages[i].p + '%';
        i++;
      }, 600);
      return iv;
    }

    const originalBranchOptions = `
  <option value="" disabled selected>Select Branch</option>
  <option value="ALL">All Branches</option>
  <option value="CE">CE</option>
  <option value="IT">IT</option>
  <option value="AIDS">AIDS</option>
  <option value="ECE">ECE</option>
  <option value="ENTC">ENTC</option>
`;

    let currentLoginRole = 'teacher';
    function switchLoginRole(role) {
      currentLoginRole = role;
      const tabTeacher = document.getElementById('tabTeacher');
      const tabAdmin = document.getElementById('tabAdmin');
      const loginTitle = document.getElementById('loginTitle');
      const loginSubtitle = document.getElementById('loginSubtitle');
      const emailLabel = document.getElementById('emailLabel');
      const emailHelp = document.getElementById('emailHelp');
      const passwordInput = document.getElementById('passwordInput');
      const emailInput = document.getElementById('emailInput');
      const err = document.getElementById('loginError');
      err.textContent = '';
      
      if (role === 'teacher') {
        tabTeacher.style.background = 'var(--surface)';
        tabTeacher.style.color = 'var(--txt-main)';
        tabTeacher.style.fontWeight = '600';
        tabTeacher.style.boxShadow = '0 1px 3px rgba(0,0,0,0.05)';
        
        tabAdmin.style.background = 'transparent';
        tabAdmin.style.color = 'var(--txt-dim)';
        tabAdmin.style.fontWeight = '500';
        tabAdmin.style.boxShadow = 'none';
        
        loginTitle.textContent = 'Faculty Login';
        loginSubtitle.textContent = 'Sign in with your @pict.edu institutional email';
        emailLabel.textContent = 'Institutional Email';
        emailInput.placeholder = 't001@pict.edu';
        emailHelp.textContent = 'Format: Teacher ID + @pict.edu (e.g. t001@pict.edu)';
        passwordInput.placeholder = 'Default password: your Teacher ID (e.g. t001)';
      } else {
        tabAdmin.style.background = 'var(--surface)';
        tabAdmin.style.color = 'var(--txt-main)';
        tabAdmin.style.fontWeight = '600';
        tabAdmin.style.boxShadow = '0 1px 3px rgba(0,0,0,0.05)';
        
        tabTeacher.style.background = 'transparent';
        tabTeacher.style.color = 'var(--txt-dim)';
        tabTeacher.style.fontWeight = '500';
        tabTeacher.style.boxShadow = 'none';
        
        loginTitle.textContent = 'Coordinator Login';
        loginSubtitle.textContent = 'Sign in to manage schedules & duties';
        emailLabel.textContent = 'Coordinator Email';
        emailInput.placeholder = 'validator@pict.edu';
        emailHelp.textContent = 'Format: admin_username@pict.edu';
        passwordInput.placeholder = 'Enter coordinator password';
      }
    }

    function handleLogin() {
      const email    = document.getElementById('emailInput').value.trim();
      const password = document.getElementById('passwordInput').value;
      const err      = document.getElementById('loginError');
      err.textContent = '';

      if (!email) {
        err.textContent = 'Please enter your institutional email.';
        return;
      }
      // Client-side domain validation
      if (!email.toLowerCase().endsWith('@pict.edu')) {
        err.textContent = 'Email must end with @pict.edu (e.g. t001@pict.edu)';
        return;
      }
      if (!password) {
        err.textContent = 'Please enter your password.';
        return;
      }

      fetch('/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, password })
      })
        .then(async r => {
          const data = await r.json();
          if (!r.ok) throw new Error(data.error || 'Login failed');
          return data;
        })
        .then(user => {
          localStorage.setItem('user', JSON.stringify(user));
          document.getElementById('passwordInput').value = '';
          showDashboard(user);
        })
        .catch(error => {
          err.textContent = error.message;
        });
    }

    function showDashboard(user) {
      if (user.is_admin) {
        document.getElementById('loginPage').style.display = 'none';
        document.getElementById('uploadPage').style.display = 'block';
        document.getElementById('teacherPage').style.display = 'none';
        document.getElementById('resultsPage').style.display = 'none';

        document.getElementById('userName').textContent = user.name;
        document.getElementById('userAvatar').textContent = user.name.charAt(0).toUpperCase();

        const confirmBtn = document.getElementById('confirmBtn');
        if (confirmBtn) {
          confirmBtn.style.display = 'inline-flex';
          confirmBtn.disabled = false;
          confirmBtn.style.background = '';
          confirmBtn.innerHTML = `
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
              <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" />
              <polyline points="22 4 12 14.01 9 11.01" />
            </svg>
            Confirm & Publish
          `;
        }

        resetConfig();
        checkConfirmedTimetable();
      } else {
        document.getElementById('loginPage').style.display = 'none';
        document.getElementById('uploadPage').style.display = 'none';
        document.getElementById('teacherPage').style.display = 'block';

        document.getElementById('teacherName').textContent = user.name;
        document.getElementById('teacherRole').textContent = user.role + " Supervisor";
        document.getElementById('teacherAvatar').textContent = user.name.charAt(0).toUpperCase();

        const upcomingBody = document.getElementById('teacherUpcomingBody');
        const pastBody     = document.getElementById('teacherPastBody');
        upcomingBody.innerHTML = '';
        pastBody.innerHTML     = '';

        const today = new Date();
        today.setHours(0, 0, 0, 0);

        const upcoming = [];
        const past     = [];

        if (user.history && user.history.length > 0) {
          user.history.forEach(d => {
            // Parse exam_date like "21-Jun-2026" or "2026-06-21"
            const examDate = parseCustomDate(d.exam_date);
            examDate.setHours(0, 0, 0, 0);
            if (!isNaN(examDate) && examDate >= today) {
              upcoming.push(d);
            } else {
              past.push(d);
            }
          });
        }

        // Sort upcoming ascending (soonest first), past descending (most recent first)
        upcoming.sort((a, b) => parseCustomDate(a.exam_date) - parseCustomDate(b.exam_date));
        past.sort((a, b) => parseCustomDate(b.exam_date) - parseCustomDate(a.exam_date));

        function buildRow(d, isUpcoming) {
          const tr = document.createElement('tr');
          tr.className = 'ri';
          const assignedDate = d.assigned_at ? new Date(d.assigned_at).toLocaleString() : 'N/A';
          
          let actionHtml = '';
          if (isUpcoming) {
            actionHtml = `
              <td>
                <button class="cta" style="padding: 6px 12px; font-size: 12px; border-radius: 8px; box-shadow: none;" 
                  onclick="openAdjustmentModal('${d.slot_id}', '${d.exam_date}', '${d.role_assigned || ''}')">
                  Request Change
                </button>
              </td>
            `;
          }
          
          const session = d.session || 'TBD';
          const courseId = d.course_id || 'ALL';
          const room = d.room_assigned || d.room || 'TBD';
          
          tr.innerHTML = `
            <td><span class="mbadge">Slot ${d.slot_id}</span></td>
            <td>${d.exam_date}</td>
            <td>${session}</td>
            <td><span class="mbadge">${courseId}</span></td>
            <td><span class="dtag">${room}</span></td>
            <td><span class="rchip ${(d.role_assigned || '').toLowerCase()}">${d.role_assigned}</span></td>
            <td>${assignedDate}</td>
            ${actionHtml}
          `;
          return tr;
        }

        if (upcoming.length > 0) {
          upcoming.forEach(d => upcomingBody.appendChild(buildRow(d, true)));
        } else {
          upcomingBody.innerHTML = '<tr><td colspan="8" style="text-align:center;color:var(--muted)">No upcoming duties scheduled.</td></tr>';
        }

        if (past.length > 0) {
          past.forEach(d => pastBody.appendChild(buildRow(d, false)));
        } else {
          pastBody.innerHTML = '<tr><td colspan="7" style="text-align:center;color:var(--muted)">No past duties on record.</td></tr>';
        }
      }
    }

    function checkConfirmedTimetable() {
      return fetch('/confirmed_timetable')
        .then(async r => {
          if (!r.ok) throw new Error('Unable to load published timetable');
          const data = await r.json();
          if (data) {
            G = data;
            render(G);
            document.getElementById('uploadPage').style.display = 'none';
            document.getElementById('resultsPage').style.display = 'block';
            
            const confirmBtn = document.getElementById('confirmBtn');
            if (confirmBtn) {
              confirmBtn.innerHTML = '<span>Published Successfully ✓</span>';
              confirmBtn.style.background = '#10b981';
              confirmBtn.disabled = true;
            }
            return true;
          }
          return false;
        })
        .catch(err => {
          console.error("Error fetching confirmed timetable:", err);
          return false;
        });
    }

    function logout() {
      localStorage.removeItem('user');
      document.getElementById('uploadPage').style.display = 'none';
      document.getElementById('teacherPage').style.display = 'none';
      document.getElementById('resultsPage').style.display = 'none';
      document.getElementById('loginPage').style.display = 'flex';
      document.getElementById('emailInput').value = '';
      document.getElementById('passwordInput').value = '';
      document.getElementById('loginError').textContent = '';
    }

    // On page load check session
    window.addEventListener('DOMContentLoaded', () => {
      const user = JSON.parse(localStorage.getItem('user'));
      if (user) {
        showDashboard(user);
      } else {
        logout();
      }
    });

    function resetConfig() {
      document.getElementById('yearSelect').value = "";
      const branchSelect = document.getElementById('branchSelect');
      branchSelect.innerHTML = originalBranchOptions;
      branchSelect.value = "";
      branchSelect.disabled = true;
      document.getElementById('examTypeSelect').value = "";
      document.getElementById('examTypeSelect').disabled = true;
      document.getElementById('slotList').classList.add('disabled');
      document.getElementById('slotList').innerHTML = '<p class="slot-placeholder">Select exam type to view slots</p>';
      document.getElementById('uploadCard').classList.add('disabled');
      document.getElementById('uzone').classList.add('disabled');
      document.getElementById('fileInput').disabled = true;
      selectedSlots.clear();
      clearFile();
    }

    function step1() {
      const year = document.getElementById('yearSelect').value;
      const branchSelect = document.getElementById('branchSelect');
      if (year === 'FY') {
        branchSelect.innerHTML = '<option value="ALL">FY (All Branches)</option>';
        branchSelect.value = 'ALL';
        branchSelect.disabled = true;
        step2();
      } else if (year) {
        branchSelect.innerHTML = originalBranchOptions;
        branchSelect.value = "";
        branchSelect.disabled = false;
      } else {
        branchSelect.innerHTML = originalBranchOptions;
        branchSelect.value = "";
        branchSelect.disabled = true;
      }
    }

    function step2() {
      const branch = document.getElementById('branchSelect').value;
      if (branch) {
        document.getElementById('examTypeSelect').disabled = false;
      }
    }

    function step3() {
      const type = document.getElementById('examTypeSelect').value;
      if (type) {
        document.getElementById('slotList').classList.remove('disabled');
        updateSlots();
        document.getElementById('uploadCard').classList.remove('disabled');
        document.getElementById('uzone').classList.remove('disabled');
        document.getElementById('fileInput').disabled = false;
      }
    }

    function reset() {
      if (G && G.confirmed_at) {
        if (!confirm("This will clear the currently published schedule from the database. Are you sure you want to create a new schedule?")) {
          return;
        }
        
        fetch('http://127.0.0.1:5000/clear_confirmed_timetable', { method: 'POST' })
          .then(r => r.json())
          .then(d => {
            if (d.success) {
              G = null;
              document.getElementById('resultsPage').style.display = 'none';
              document.getElementById('uploadPage').style.display = 'block';
              resetConfig();
            } else {
              alert("Error clearing confirmed schedule: " + d.error);
            }
          })
          .catch(err => {
            alert("Connection error clearing schedule.");
            console.error(err);
          });
      } else {
        G = null;
        document.getElementById('resultsPage').style.display = 'none';
        document.getElementById('uploadPage').style.display = 'block';
        resetConfig();
      }
    }

    const selectedSlots = new Set();
    function updateSlots() {
      const type = document.getElementById('examTypeSelect').value;
      const list = document.getElementById('slotList');
      const slots = type === 'Insem'
        ? ['10:00 – 11:00', '14:00 – 15:00', '16:00 – 17:00']
        : type === 'Endsem'
          ? ['10:00 – 12:30', '14:00 – 16:30']
          : ['12:00 – 14:00', '14:30 – 16:30', '17:00 – 19:00']; // Practical slots

      list.innerHTML = slots.map(s => `
    <div class="slot-item ${selectedSlots.has(s) ? 'active' : ''}" onclick="toggleSlot('${s}', this)">
      ${s}
    </div>
  `).join('');
    }

    function toggleSlot(s, el) {
      if (selectedSlots.has(s)) {
        selectedSlots.delete(s);
        el.classList.remove('active');
      } else {
        selectedSlots.add(s);
        el.classList.add('active');
      }
    }

    let G = null;
    // exam dates are provided in the uploaded Excel; no client-side start date
    function generate() {
      const f = fi.files[0]; if (!f) return;
      const year = document.getElementById('yearSelect').value;
      const branch = document.getElementById('branchSelect').value;
      const examType = document.getElementById('examTypeSelect').value;

      document.getElementById('loadScreen').style.display = 'flex';
      const iv = startLoad();

      const fd = new FormData();
      fd.append('file', f);
      fd.append('year', year);
      fd.append('branch', branch);
      fd.append('exam_type', examType);
      // Note: selected slots/date are handled server-side now; do not send start_date
      fd.append('selected_slots', Array.from(selectedSlots).join(','));

      fetch('http://127.0.0.1:5000/generate', { method: 'POST', body: fd })
        .then(async r => {
          const data = await r.json();
          if (!r.ok) throw new Error(data.error || 'Server error');
          return data;
        })
        .then(d => {
          G = d;
          
          // Reset confirm button state for newly generated timetable
          const confirmBtn = document.getElementById('confirmBtn');
          if (confirmBtn) {
            confirmBtn.disabled = false;
            confirmBtn.style.background = '';
            confirmBtn.innerHTML = `
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
                <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" />
                <polyline points="22 4 12 14.01 9 11.01" />
              </svg>
              Confirm & Publish
            `;
          }

          // Dates come from the Excel upload; no client-side start date stored
          setTimeout(() => {
            render(d);
            document.getElementById('loadScreen').style.display = 'none';
            document.getElementById('uploadPage').style.display = 'none';
            document.getElementById('resultsPage').style.display = 'block';
          }, 1500);
        })
        .catch(err => {
          document.getElementById('loadScreen').style.display = 'none';
          alert('Generation Failed: ' + err.message);
        });
    }

    const slotCls = ['', 's1c', 's2c', 's3c', 's4c', 's5c', 's6c'];
    function schip(s) { return `<span class="schip ${slotCls[s] || 's1c'}">Slot ${s}</span>` }

    function render(d) {
      const s = d.summary;
      document.getElementById('sc1').textContent = s.total_courses;
      document.getElementById('sc2').textContent = s.total_students;
      document.getElementById('sc3').textContent = s.slots_used;
      document.getElementById('sc4').textContent = s.conflicts_found === 0 ? '0 ✓' : s.conflicts_found;
      document.getElementById('sc5').textContent = s.duties_assigned + '/' + (s.duties_assigned + s.duties_failed);
      document.getElementById('sc6').textContent = s.pref_satisfaction + '%';
      document.getElementById('rsub').textContent =
        s.total_courses + ' courses · ' + s.slots_used + ' slots · ' + s.total_students + ' students';

      const tb = document.querySelector('#ttbl tbody'); tb.innerHTML = '';
      d.timetable.forEach((r, i) => {
        const tr = document.createElement('tr');
        tr.className = 'ri'; tr.style.animationDelay = (i * 28) + 'ms';
        tr.innerHTML = `
      <td><span class="mbadge">${r.course_id}</span></td>
      <td class="ncell">${r.course_name}</td>
      <td>${r.year}</td>
      <td><span class="dtag">${r.department}</span></td>
      <td>${schip(r.slot)}</td>
      <td class="mcell">${r.date}</td>
      <td>${r.session}</td>
      <td><strong>${r.enrolled_students}</strong></td>
      <td class="mcell">${r.room_assigned || r['Room Assigned'] || ''}</td>`;
        tb.appendChild(tr);
      });

      const rb = document.querySelector('#rtbl tbody'); rb.innerHTML = '';
      d.room_allocation.forEach((r, i) => {
        const tr = document.createElement('tr');
        tr.className = 'ri'; tr.style.animationDelay = (i * 28) + 'ms';
        tr.innerHTML = `
      <td><span class="mbadge">${r.course_id}</span></td>
      <td>${schip(r.slot)}</td>
      <td class="mcell">${r.date}</td>
      <td>${r.session}</td>
      <td class="mcell rcell">${r.rooms_assigned}</td>
      <td>${r.total_capacity}</td>
      <td><span class="okchip">✓ ${r.status}</span></td>`;
        rb.appendChild(tr);
      });

      const tc = document.querySelector('#tctbl tbody'); tc.innerHTML = '';
      d.teacher_duties.forEach((r, i) => {
        const tr = document.createElement('tr');
        tr.className = 'ri'; tr.style.animationDelay = (i * 28) + 'ms';
        const rl = r.role.toLowerCase();
        tr.innerHTML = `
      <td>${schip(r.slot)}</td>
      <td class="mcell">${r.date}</td>
      <td>${r.session}</td>
      <td><span class="mbadge">${r.course_id}</span></td>
      <td><span class="rchip ${rl}">${r.role}</span></td>
      <td class="mcell">${r.teacher_id}</td>
      <td>
        <div class="t-info">
          <span class="t-name">${r.teacher_name}</span>
        </div>
      </td>
      <td><span class="hbadge">${r.last_role}</span></td>
      <td>${(r.is_priority === 'Yes' || r.is_priority === true) ? '<span class="pyes">✓ Yes</span>' : '<span class="pno">— No</span>'}</td>
      <td class="mcell">${r.room_assigned || r['Room Assigned'] || ''}</td>`;
        tc.appendChild(tr);
      });

      document.getElementById('uploadPage').style.display = 'none';
      document.getElementById('resultsPage').style.display = 'block';
      setTimeout(() => moveInd(document.querySelector('.ptab.on')), 80);
    }

    function tab(name, btn) {
      document.querySelectorAll('.tcont').forEach(t => t.classList.remove('on'));
      document.querySelectorAll('.ptab').forEach(t => t.classList.remove('on'));
      document.getElementById('tc-' + name).classList.add('on');
      btn.classList.add('on');
      moveInd(btn);
    }
    function moveInd(btn) {
      if (!btn) return;
      const ind = document.getElementById('pind');
      ind.style.left = btn.offsetLeft + 'px';
      ind.style.width = btn.offsetWidth + 'px';
    }
    window.addEventListener('resize', () => moveInd(document.querySelector('.ptab.on')));

    function resolveDate(dayLabel) {
      // Convert "Day N" to an actual calendar date using the start_date stored globally
      if (!window._examStartDate) return dayLabel;
      const match = String(dayLabel).match(/Day\s*(\d+)/i);
      if (!match) return dayLabel; // already a real date string, return as-is
      const dayOffset = parseInt(match[1], 10) - 1;
      const base = new Date(window._examStartDate);
      base.setDate(base.getDate() + dayOffset);
      return base.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
    }

    function exportCSV(type) {
      if (!G) return;
      let rows = [], fn = '';
      if (type === 'tt') {
        fn = 'exam_timetable.csv';
        rows = [['course_id', 'course_name', 'year', 'department', 'slot', 'date', 'enrolled_students', 'room_assigned'],
        ...G.timetable.map(r => [
          r.course_id,
          `"${r.course_name}"`,
          r.year,
          r.department,
          r.slot,
          resolveDate(r.date),
          r.enrolled_students,
          (r.room_assigned || r['Room Assigned'] || '')
        ])];
      } else if (type === 'rm') {
        fn = 'room_allocation.csv';
        rows = [['course_id', 'slot', 'date', 'rooms_assigned', 'total_capacity', 'status'],
        ...G.room_allocation.map(r => [
          r.course_id,
          r.slot,
          resolveDate(r.date),
          `"${r.rooms_assigned}"`,
          r.total_capacity,
          r.status
        ])];
      } else {
        fn = 'teacher_duties.csv';
        rows = [['slot', 'date', 'course_id', 'role', 'teacher_id', 'teacher_name', 'is_priority', 'room_assigned'],
        ...G.teacher_duties.map(r => [
          r.slot,
          resolveDate(r.date),
          r.course_id,
          r.role,
          r.teacher_id,
          `"${r.teacher_name}"`,
          (r.is_priority === 'Yes' || r.is_priority === true) ? 'Yes' : 'No',
          (r.room_assigned || r['Room Assigned'] || '')
        ])];
      }
      const csvContent = rows.map(r => r.join(',')).join('\n');
      const a = Object.assign(document.createElement('a'), {
        href: URL.createObjectURL(new Blob(['﻿' + csvContent], { type: 'text/csv;charset=utf-8;' })),
        download: fn
      });
      a.click(); URL.revokeObjectURL(a.href);
    }

    function exportPDF(type) {
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF('l', 'mm', 'a4');

      const titles = { tt: 'Exam Timetable', rm: 'Room Allocation', tc: 'Teacher Duties' };
      const tables = { tt: '#ttbl', rm: '#rtbl', tc: '#tctbl' };

      doc.setFont("Inter", "bold");
      doc.setFontSize(22);
      doc.setTextColor(30, 41, 59);
      doc.text(titles[type], 14, 20);

      doc.setFontSize(10);
      doc.setFont("Inter", "normal");
      doc.setTextColor(100);
      doc.text(`Generated on ${new Date().toLocaleString()} | ExamSched System`, 14, 28);

      doc.autoTable({
        html: tables[type],
        startY: 35,
        margin: { top: 35, left: 14, right: 14, bottom: 20 },
        styles: {
          fontSize: 8,
          cellPadding: 4,
          font: "helvetica",
          textColor: [51, 65, 85],
          lineColor: [226, 232, 240],
          lineWidth: 0.1
        },
        headStyles: {
          fillColor: [99, 102, 241],
          textColor: [255, 255, 255],
          fontStyle: 'bold',
          halign: 'left'
        },
        alternateRowStyles: {
          fillColor: [248, 250, 252]
        },
        didDrawPage: function (data) {
          // Footer
          doc.setFontSize(8);
          doc.text(`Page ${doc.internal.getNumberOfPages()}`, data.settings.margin.left, doc.internal.pageSize.height - 10);
        }
      });

      doc.save(`${titles[type].toLowerCase().replace(/ /g, '_')}_${new Date().getTime()}.pdf`);
    }

    function downloadTeacherDutiesPDF(teacherId) {
      fetch(`/api/teacher/duties?teacher_id=${encodeURIComponent(teacherId)}`)
        .then(async r => {
          const data = await r.json();
          if (!r.ok) throw new Error(data.error || 'Failed to fetch teacher duties.');
          return data;
        })
        .then(teacher => {
          const { jsPDF } = window.jspdf;
          const doc = new jsPDF('l', 'mm', 'a4');

          doc.setFont("Inter", "bold");
          doc.setFontSize(22);
          doc.setTextColor(30, 41, 59);
          doc.text(`Teacher Duties — ${teacher.name}`, 14, 20);

          doc.setFontSize(10);
          doc.setFont("Inter", "normal");
          doc.setTextColor(100);
          doc.text(`Generated on ${new Date().toLocaleString()} | ExamSched System`, 14, 28);

          const rows = (teacher.history || []).map(h => {
            const slotText = `Slot ${h.slot_id || h.slot || ''}`;
            const dateText = h.exam_date || '';
            const sessionText = h.session || 'TBD';
            const courseText = h.course_id || 'ALL';
            const roleText = h.role_assigned || h.role || 'Junior';
            const teacherIdText = teacher.teacher_id || '';
            const teacherNameText = teacher.name || '';
            const historyText = teacher.last_role || 'N/A';
            const priorityText = (teacher.has_served_high_role === false || teacher.has_served_high_role === 'No') ? '✓ Yes' : '— No';
            const roomText = h.room_assigned || h.room || 'TBD';

            return [
              slotText,
              dateText,
              sessionText,
              courseText,
              roleText,
              teacherIdText,
              teacherNameText,
              historyText,
              priorityText,
              roomText
            ];
          });

          // Sort rows chronologically by date and slot
          rows.sort((a, b) => {
            const dateA = parseCustomDate(a[1]);
            const dateB = parseCustomDate(b[1]);
            if (dateA - dateB !== 0) return dateA - dateB;
            const slotA = parseInt(a[0].replace('Slot ', ''), 10) || 0;
            const slotB = parseInt(b[0].replace('Slot ', ''), 10) || 0;
            return slotA - slotB;
          });

          doc.autoTable({
            head: [['Slot', 'Date', 'Session', 'Course', 'Role', 'Teacher ID', 'Teacher', 'History', 'Priority', 'Room']],
            body: rows,
            startY: 35,
            margin: { top: 35, left: 14, right: 14, bottom: 20 },
            styles: {
              fontSize: 8,
              cellPadding: 4,
              font: "helvetica",
              textColor: [51, 65, 85],
              lineColor: [226, 232, 240],
              lineWidth: 0.1
            },
            headStyles: {
              fillColor: [99, 102, 241],
              textColor: [255, 255, 255],
              fontStyle: 'bold',
              halign: 'left'
            },
            alternateRowStyles: {
              fillColor: [248, 250, 252]
            },
            didDrawPage: function (data) {
              doc.setFontSize(8);
              doc.text(`Page ${doc.internal.getNumberOfPages()}`, data.settings.margin.left, doc.internal.pageSize.height - 10);
            }
          });

          const safeName = teacher.name.toLowerCase().replace(/ /g, '_');
          doc.save(`teacher_duties_${safeName}_${new Date().getTime()}.pdf`);
        })
        .catch(err => {
          alert('Failed to generate PDF: ' + err.message);
        });
    }

    function downloadAllTeacherDutiesPDF() {
      fetch('/api/timetable/duties')
        .then(async r => {
          const data = await r.json();
          if (!r.ok) throw new Error(data.error || 'Failed to fetch duties.');
          return data;
        })
        .then(payload => {
          const duties = payload.duties || [];
          // Group by teacher_id
          const byTeacher = {};
          duties.forEach(d => {
            const tid = d.teacher_id || 'unknown';
            if (!byTeacher[tid]) byTeacher[tid] = { teacher_id: tid, teacher_name: d.teacher_name || '', history: [] };
            byTeacher[tid].history.push({
              slot_id: d.slot,
              exam_date: d.date,
              session: d.session,
              course_id: d.course_id,
              role_assigned: d.role,
              room_assigned: d.room_assigned
            });
          });

          const { jsPDF } = window.jspdf;
          const doc = new jsPDF('l', 'mm', 'a4');
          let firstPage = true;

          Object.values(byTeacher).forEach((teacher, idx) => {
            if (!firstPage) doc.addPage();
            firstPage = false;

            doc.setFont("Inter", "bold");
            doc.setFontSize(20);
            doc.setTextColor(30, 41, 59);
            const title = `Teacher Duties — ${teacher.teacher_name || teacher.teacher_id}`;
            doc.text(title, 14, 20);

            doc.setFontSize(9);
            doc.setFont("Inter", "normal");
            doc.setTextColor(100);
            doc.text(`Generated on ${new Date().toLocaleString()} | ExamSched System`, 14, 28);

            const rows = (teacher.history || []).map(h => {
              return [
                `Slot ${h.slot_id}`,
                h.exam_date || '',
                h.session || 'TBD',
                h.course_id || 'ALL',
                h.role_assigned || 'Junior',
                teacher.teacher_id || '',
                teacher.teacher_name || '',
                h.room_assigned || 'TBD'
              ];
            });

            // Sort rows by date then slot
            rows.sort((a, b) => {
              const da = parseCustomDate(a[1]);
              const db = parseCustomDate(b[1]);
              if (da - db !== 0) return da - db;
              const sa = parseInt(a[0].replace('Slot ', ''), 10) || 0;
              const sb = parseInt(b[0].replace('Slot ', ''), 10) || 0;
              return sa - sb;
            });

            doc.autoTable({
              head: [['Slot', 'Date', 'Session', 'Course', 'Role', 'Teacher ID', 'Teacher', 'Room']],
              body: rows,
              startY: 35,
              margin: { top: 35, left: 14, right: 14, bottom: 20 },
              styles: {
                fontSize: 8,
                cellPadding: 4,
                font: "helvetica",
                textColor: [51, 65, 85],
                lineColor: [226, 232, 240],
                lineWidth: 0.1
              },
              headStyles: {
                fillColor: [99, 102, 241],
                textColor: [255, 255, 255],
                fontStyle: 'bold',
                halign: 'left'
              },
              alternateRowStyles: { fillColor: [248, 250, 252] },
              didDrawPage: function (data) {
                doc.setFontSize(8);
                doc.text(`Page ${doc.internal.getNumberOfPages()}`, data.settings.margin.left, doc.internal.pageSize.height - 10);
              }
            });
          });

          doc.save(`all_teacher_duties_${new Date().getTime()}.pdf`);
        })
        .catch(err => {
          alert('Failed to generate PDF: ' + err.message);
        });
    }

    function confirmSchedule() {
      if (!G) return;
      const btn = document.getElementById('confirmBtn');
      const oldText = btn.innerHTML;
      btn.disabled = true;
      btn.innerHTML = '<span>Publishing...</span>';

      fetch('/confirm', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          timetable: G.timetable,
          room_allocation: G.room_allocation,
          assignments: G.teacher_duties,
          summary: G.summary,
          department: document.getElementById('branchSelect').value
        })
      })
        .then(r => r.json())
        .then(d => {
          if (d.success) {
            alert('Schedule confirmed and faculty history updated!' + (d.reset_triggered ? '\n\n🔄 Fairness cycle complete. All flags have been reset.' : ''));
            btn.innerHTML = '<span>Published Successfully ✓</span>';
            btn.style.background = '#10b981';
          } else {
            alert('Error: ' + d.error);
            btn.disabled = false;
            btn.innerHTML = oldText;
          }
        })
        .catch(err => {
          alert('Connection error');
          btn.disabled = false;
          btn.innerHTML = oldText;
        });
    }


    setTimeout(() => moveInd(document.querySelector('.ptab.on')), 150);

    function submitChangePassword() {
      const current = document.getElementById('cpCurrent').value;
      const newPw   = document.getElementById('cpNew').value;
      const confirm = document.getElementById('cpConfirm').value;
      const msg     = document.getElementById('cpMsg');

      msg.textContent = '';
      msg.style.color = '';

      if (!current || !newPw || !confirm) {
        msg.textContent = 'All fields are required.';
        msg.style.color = '#f87171';
        return;
      }
      if (newPw.length < 6) {
        msg.textContent = 'New password must be at least 6 characters.';
        msg.style.color = '#f87171';
        return;
      }
      if (newPw !== confirm) {
        msg.textContent = 'New passwords do not match.';
        msg.style.color = '#f87171';
        return;
      }

      const user = JSON.parse(localStorage.getItem('user') || '{}');
      const identifier = user.identifier || '';
      if (!identifier) {
        msg.textContent = 'Session error — please log out and log in again.';
        msg.style.color = '#f87171';
        return;
      }

      fetch('/change_password', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ identifier, current_password: current, new_password: newPw })
      })
        .then(r => r.json())
        .then(d => {
          if (d.success) {
            msg.textContent = '✓ Password updated successfully.';
            msg.style.color = '#10b981';
            document.getElementById('cpCurrent').value = '';
            document.getElementById('cpNew').value = '';
            document.getElementById('cpConfirm').value = '';
          } else {
            msg.textContent = d.error || 'Failed to update password.';
            msg.style.color = '#f87171';
          }
        })
        .catch(() => {
          msg.textContent = 'Connection error. Please try again.';
          msg.style.color = '#f87171';
        });
    }

    function openAdjustmentModal(slot_id, exam_date, role) {
      document.getElementById('adjSlot').value = slot_id;
      document.getElementById('adjDate').value = exam_date;
      document.getElementById('adjRole').value = role;
      
      document.getElementById('modalDutyDetails').innerHTML = `
        <strong>Date:</strong> ${exam_date} &nbsp;|&nbsp; 
        <strong>Slot:</strong> ${slot_id} &nbsp;|&nbsp; 
        <strong>Role:</strong> ${role}
      `;
      
      document.getElementById('modalFile').value = '';
      document.getElementById('modalFileStatus').textContent = 'Drag & drop or click to browse';
      document.getElementById('adjReason').value = '';
      document.getElementById('modalError').textContent = '';
      
      document.getElementById('adjustmentModal').style.display = 'flex';
    }

    function closeAdjustmentModal() {
      document.getElementById('adjustmentModal').style.display = 'none';
    }

    function handleModalFileSelect(input) {
      const f = input.files[0];
      if (f) {
        document.getElementById('modalFileStatus').textContent = f.name + ' (' + (f.size / 1024 / 1024).toFixed(2) + ' MB)';
      }
    }

    function submitAdjustmentRequest(e) {
      e.preventDefault();
      const err = document.getElementById('modalError');
      err.textContent = '';
      
      const fileInput = document.getElementById('modalFile');
      const f = fileInput.files[0];
      if (!f) {
        err.textContent = 'Please select an application file to upload.';
        return;
      }
      
      const user = JSON.parse(localStorage.getItem('user') || '{}');
      const teacher_id = user.identifier || '';
      
      const fd = new FormData();
      fd.append('file', f);
      fd.append('teacher_id', teacher_id);
      fd.append('current_date', document.getElementById('adjDate').value);
      fd.append('current_slot', document.getElementById('adjSlot').value);
      fd.append('current_session', document.getElementById('adjRole').value === 'Senior' ? 'Morning' : 'Afternoon');
      fd.append('reason', document.getElementById('adjReason').value.trim());
      
      const btn = e.target.querySelector('button[type="submit"]');
      const originalText = btn.innerHTML;
      btn.disabled = true;
      btn.innerHTML = '<span>Submitting...</span>';
      
      fetch('/request_adjustment', {
        method: 'POST',
        body: fd
      })
      .then(async r => {
        const d = await r.json();
        if (!r.ok) throw new Error(d.error || 'Submission failed');
        return d;
      })
      .then(d => {
        alert(d.message);
        closeAdjustmentModal();
        refreshTeacherDuties(teacher_id);
      })
      .catch(error => {
        err.textContent = error.message;
      })
      .finally(() => {
        btn.disabled = false;
        btn.innerHTML = originalText;
      });
    }

    function refreshTeacherDuties(identifier) {
      fetch('/teacher/status?identifier=' + encodeURIComponent(identifier))
      .then(r => r.json())
      .then(user => {
        localStorage.setItem('user', JSON.stringify(user));
        showDashboard(user);
      })
      .catch(err => {
        console.error('Error refreshing teacher dashboard:', err);
      });
    }

    let activeRescheduleTab = 'swap';
    let currentAlternatives = null;

    function showAdminSection(section) {
      const navGen = document.getElementById('navGenerator');
      const navAdj = document.getElementById('navAdjustments');
      const genView = document.getElementById('generatorView');
      const adjView = document.getElementById('adjustmentsView');

      if (section === 'generator') {
        navGen.classList.add('active');
        navAdj.classList.remove('active');
        genView.style.display = 'grid';
        adjView.style.display = 'none';
      } else {
        navGen.classList.remove('active');
        navAdj.classList.add('active');
        genView.style.display = 'none';
        adjView.style.display = 'block';
        loadAdjustments();
      }
    }

    function loadAdjustments() {
      const tbody = document.getElementById('adjustmentsTableBody');
      tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; color:var(--txt-dim);">Loading adjustment requests...</td></tr>';

      fetch('/adjustments')
        .then(r => r.json())
        .then(data => {
          tbody.innerHTML = '';
          if (data.length === 0) {
            tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; color:var(--txt-dim);">No change requests found.</td></tr>';
            return;
          }
          data.forEach(r => {
            const tr = document.createElement('tr');
            tr.className = 'ri';
            
            const fileLink = r.file_path ? `<a href="/uploads/${r.file_path}" target="_blank" class="ulink" style="font-weight:600;">View File 📄</a>` : 'No file';
            const commentsText = r.comments ? `<br><small style="color:var(--txt-dim)">Note: ${r.comments}</small>` : '';
            const statusChip = `<span class="stag" style="background:${r.status === 'Pending' ? 'rgba(245,158,11,.1)' : r.status === 'Approved' ? 'rgba(16,185,129,.1)' : 'rgba(239,68,68,.1)'}; color:${r.status === 'Pending' ? '#d97706' : r.status === 'Approved' ? '#10b981' : '#ef4444'}; font-weight:700;">${r.status}</span>${commentsText}`;
            
            let actionHtml = '—';
            if (r.status === 'Pending') {
              actionHtml = `
                <div style="display:flex; gap:8px;">
                  <button class="cta" style="padding:6px 12px; font-size:12px; border-radius:8px; box-shadow:none; background:var(--acc-ind);" onclick="openCoordinatorApprovalModal('${r._id}', '${r.teacher_id}', '${r.teacher_name}', '${r.current_date}', '${r.current_slot}', '${r.current_session}')">
                    Approve
                  </button>
                  <button class="cta" style="padding:6px 12px; font-size:12px; border-radius:8px; box-shadow:none; background:#ef4444;" onclick="openCoordinatorRejectionModal('${r._id}')">
                    Reject
                  </button>
                </div>
              `;
            }

            let newDutiesHtml = '—';
            if (r.status === 'Approved') {
              newDutiesHtml = `
                <button class="cta" style="padding:6px 12px; font-size:12px; border-radius:8px; box-shadow:none; background:linear-gradient(135deg, var(--acc-ind), var(--acc-vio));" onclick="downloadAllTeacherDutiesPDF()">
                  Download Duties 📄
                </button>
              `;
            }
            
            tr.innerHTML = `
              <td>
                <strong>${r.teacher_name}</strong><br>
                <small style="color:var(--txt-dim); font-family:var(--mono);">${r.teacher_id}</small>
              </td>
              <td>
                <span class="mbadge">Slot ${r.current_slot}</span><br>
                <small style="color:var(--txt-dim); font-family:var(--mono);">${r.current_date}</small>
              </td>
              <td style="max-width: 250px; white-space: normal;">
                <span style="font-style:italic; font-size:14px; color:var(--txt-main);">"${r.reason || 'No reason provided'}"</span><br>
                ${fileLink}
              </td>
              <td>${statusChip}</td>
              <td>${actionHtml}</td>
              <td>${newDutiesHtml}</td>
            `;
            tbody.appendChild(tr);
          });
        })
        .catch(err => {
          tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; color:#ef4444;">Error loading adjustments: ' + err.message + '</td></tr>';
        });
    }

    function openCoordinatorApprovalModal(reqId, tId, tName, date, slot, session) {
      document.getElementById('coordRequestID').value = reqId;
      document.getElementById('coordinatorModalDutyDetails').innerHTML = `
        <strong>Teacher:</strong> ${tName} (${tId})<br>
        <strong>Current Duty:</strong> ${date} | Slot ${slot} (${session})
      `;

      document.getElementById('swapPartnerSelect').innerHTML = '<option value="" disabled selected>Loading recommended swaps...</option>';
      document.getElementById('moveSlotSelect').innerHTML = '<option value="" disabled selected>Loading conflict-free slots...</option>';
      document.getElementById('swapPreferenceMatchNote').textContent = '';
      document.getElementById('coordinatorModalError').textContent = '';
      
      switchRescheduleTab('swap');
      document.getElementById('coordinatorApprovalModal').style.display = 'flex';

      fetch(`/adjustments/alternatives?request_id=${reqId}`)
        .then(r => r.json())
        .then(d => {
          currentAlternatives = d;
          populateRescheduleOptions();
        })
        .catch(err => {
          document.getElementById('coordinatorModalError').textContent = 'Failed to load alternatives: ' + err.message;
        });
    }

    function closeCoordinatorApprovalModal() {
      document.getElementById('coordinatorApprovalModal').style.display = 'none';
    }

    function switchRescheduleTab(tab) {
      activeRescheduleTab = tab;
      const tSwap = document.getElementById('rescheduleTabSwap');
      const tMove = document.getElementById('rescheduleTabMove');
      const swapSelect = document.getElementById('swapSelectorContainer');
      const moveSelect = document.getElementById('moveSelectorContainer');

      if (tab === 'swap') {
        tSwap.style.background = 'var(--surface)';
        tSwap.style.color = 'var(--txt-main)';
        tSwap.style.fontWeight = '600';
        tSwap.style.boxShadow = '0 1px 3px rgba(0,0,0,0.05)';
        
        tMove.style.background = 'transparent';
        tMove.style.color = 'var(--txt-dim)';
        tMove.style.fontWeight = '500';
        tMove.style.boxShadow = 'none';

        swapSelect.style.display = 'block';
        moveSelect.style.display = 'none';
      } else {
        tMove.style.background = 'var(--surface)';
        tMove.style.color = 'var(--txt-main)';
        tMove.style.fontWeight = '600';
        tMove.style.boxShadow = '0 1px 3px rgba(0,0,0,0.05)';
        
        tSwap.style.background = 'transparent';
        tSwap.style.color = 'var(--txt-dim)';
        tSwap.style.fontWeight = '500';
        tSwap.style.boxShadow = 'none';

        swapSelect.style.display = 'none';
        moveSelect.style.display = 'block';
      }
    }

    function populateRescheduleOptions() {
      if (!currentAlternatives) return;

      const swapSelect = document.getElementById('swapPartnerSelect');
      swapSelect.innerHTML = '<option value="" disabled selected>Select a swap partner</option>';
      if (currentAlternatives.swap_options.length === 0) {
        swapSelect.innerHTML = '<option value="" disabled>No compatible swap partners found</option>';
      } else {
        currentAlternatives.swap_options.forEach((opt, idx) => {
          const matchNote = opt.pref_score === 2 ? '★ Both prefer' : opt.pref_score === 1 ? '✦ One prefers' : 'No preference match';
          const optionText = `${opt.teacher_name} - ${opt.date} (Slot ${opt.slot}) [${matchNote}]`;
          const optionEl = document.createElement('option');
          optionEl.value = idx;
          optionEl.textContent = optionText;
          swapSelect.appendChild(optionEl);
        });
      }

      const moveSelect = document.getElementById('moveSlotSelect');
      moveSelect.innerHTML = '<option value="" disabled selected>Select an empty slot</option>';
      if (currentAlternatives.free_slots.length === 0) {
        moveSelect.innerHTML = '<option value="" disabled>No empty slots found</option>';
      } else {
        currentAlternatives.free_slots.forEach((opt, idx) => {
          const optionText = `${opt.date} (Slot ${opt.slot})`;
          const optionEl = document.createElement('option');
          optionEl.value = idx;
          optionEl.textContent = optionText;
          moveSelect.appendChild(optionEl);
        });
      }
    }

    function handleSwapPartnerSelectChange(select) {
      if (!currentAlternatives || select.value === "") return;
      const opt = currentAlternatives.swap_options[parseInt(select.value, 10)];
      const note = document.getElementById('swapPreferenceMatchNote');
      
      let text = '';
      if (opt.requester_preferred && opt.partner_preferred) {
        text = '✓ Perfect Match! Requester prefers Target Slot, and Partner prefers Current Slot.';
      } else if (opt.requester_preferred) {
        text = '✦ Partial Match: Requester prefers Target Slot.';
      } else if (opt.partner_preferred) {
        text = '✦ Partial Match: Swap partner prefers Current Slot.';
      } else {
        text = 'No preference match, but slots are free of clashes.';
      }
      note.textContent = text;
    }

    function submitCoordinatorApproval(e) {
      e.preventDefault();
      const err = document.getElementById('coordinatorModalError');
      err.textContent = '';

      const reqId = document.getElementById('coordRequestID').value;
      const type = activeRescheduleTab;
      
      let payload = { request_id: reqId, type: type };
      
      if (type === 'swap') {
        const select = document.getElementById('swapPartnerSelect');
        if (select.value === "") {
          err.textContent = 'Please select a swap partner.';
          return;
        }
        const opt = currentAlternatives.swap_options[parseInt(select.value, 10)];
        payload.swap_teacher_id = opt.teacher_id;
        payload.new_date = opt.date;
        payload.new_slot = opt.slot;
      } else {
        const select = document.getElementById('moveSlotSelect');
        if (select.value === "") {
          err.textContent = 'Please select an empty slot.';
          return;
        }
        const opt = currentAlternatives.free_slots[parseInt(select.value, 10)];
        payload.new_date = opt.date;
        payload.new_slot = opt.slot;
      }

      const btn = e.target.querySelector('button[type="submit"]');
      const originalText = btn.innerHTML;
      btn.disabled = true;
      btn.innerHTML = '<span>Processing...</span>';

      fetch('/approve_adjustment', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      })
      .then(async r => {
        const d = await r.json();
        if (!r.ok) throw new Error(d.error || 'Rescheduling failed');
        return d;
      })
      .then(d => {
        alert(d.message);
        closeCoordinatorApprovalModal();
        loadAdjustments();
      })
      .catch(error => {
        err.textContent = error.message;
      })
      .finally(() => {
        btn.disabled = false;
        btn.innerHTML = originalText;
      });
    }

    function openCoordinatorRejectionModal(reqId) {
      document.getElementById('coordRejectRequestID').value = reqId;
      document.getElementById('coordRejectComments').value = '';
      document.getElementById('coordinatorRejectModalError').textContent = '';
      document.getElementById('coordinatorRejectionModal').style.display = 'flex';
    }

    function closeCoordinatorRejectionModal() {
      document.getElementById('coordinatorRejectionModal').style.display = 'none';
    }

    function submitCoordinatorRejection(e) {
      e.preventDefault();
      const err = document.getElementById('coordinatorRejectModalError');
      err.textContent = '';

      const reqId = document.getElementById('coordRejectRequestID').value;
      const comments = document.getElementById('coordRejectComments').value.trim();

      const btn = e.target.querySelector('button[type="submit"]');
      const originalText = btn.innerHTML;
      btn.disabled = true;
      btn.innerHTML = '<span>Rejecting...</span>';

      fetch('/reject_adjustment', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ request_id: reqId, comments: comments })
      })
      .then(async r => {
        const d = await r.json();
        if (!r.ok) throw new Error(d.error || 'Rejection failed');
        return d;
      })
      .then(d => {
        alert(d.message);
        closeCoordinatorRejectionModal();
        loadAdjustments();
      })
      .catch(error => {
        err.textContent = error.message;
      })
      .finally(() => {
        btn.disabled = false;
        btn.innerHTML = originalText;
      });
    }

    window.addEventListener('DOMContentLoaded', () => {
      const modalUzone = document.getElementById('modalUzone');
      if (modalUzone) {
        modalUzone.addEventListener('dragover', e => {
          e.preventDefault();
          modalUzone.style.borderColor = 'var(--acc-ind)';
          modalUzone.style.background = 'rgba(99, 102, 241, 0.02)';
        });
        modalUzone.addEventListener('dragleave', () => {
          modalUzone.style.borderColor = '#e2e8f0';
          modalUzone.style.background = '#fbfcfd';
        });
        modalUzone.addEventListener('drop', e => {
          e.preventDefault();
          modalUzone.style.borderColor = '#e2e8f0';
          modalUzone.style.background = '#fbfcfd';
          const fileInput = document.getElementById('modalFile');
          const f = e.dataTransfer.files[0];
          if (f) {
            const container = new DataTransfer();
            container.items.add(f);
            fileInput.files = container.files;
            handleModalFileSelect(fileInput);
          }
        });
      }
    });
  