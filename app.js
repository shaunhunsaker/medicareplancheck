// ===== CONFETTI =====
function launchConfetti() {
  const canvas = document.getElementById('confettiCanvas');
  const ctx = canvas.getContext('2d');
  canvas.width  = window.innerWidth;
  canvas.height = window.innerHeight;

  const colors = ['#1a56db','#0a8f5c','#f59e0b','#e11d48','#7dd3fc','#4ade80'];
  const pieces = Array.from({ length: 120 }, () => ({
    x:  Math.random() * canvas.width,
    y:  Math.random() * -canvas.height,
    w:  6 + Math.random() * 8,
    h:  10 + Math.random() * 6,
    color: colors[Math.floor(Math.random() * colors.length)],
    rot: Math.random() * 360,
    vx: -2 + Math.random() * 4,
    vy: 3 + Math.random() * 4,
    vr: -3 + Math.random() * 6,
  }));

  let frame;
  let elapsed = 0;
  function draw() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    pieces.forEach(p => {
      ctx.save();
      ctx.translate(p.x + p.w / 2, p.y + p.h / 2);
      ctx.rotate(p.rot * Math.PI / 180);
      ctx.fillStyle = p.color;
      ctx.fillRect(-p.w / 2, -p.h / 2, p.w, p.h);
      ctx.restore();
      p.x  += p.vx;
      p.y  += p.vy;
      p.rot += p.vr;
    });
    elapsed++;
    if (elapsed < 180) frame = requestAnimationFrame(draw);
    else { ctx.clearRect(0, 0, canvas.width, canvas.height); }
  }
  draw();
}

// Sticky header shadow on scroll
const header = document.getElementById('siteHeader');
window.addEventListener('scroll', () => {
  header.classList.toggle('scrolled', window.scrollY > 10);
}, { passive: true });

// FAQ: collapse others when one opens
document.querySelectorAll('.faq-item').forEach(item => {
  item.addEventListener('toggle', () => {
    if (item.open) {
      document.querySelectorAll('.faq-item[open]').forEach(other => {
        if (other !== item) other.removeAttribute('open');
      });
    }
  });
});

// Click-to-call tracking
document.querySelectorAll('a[href^="tel:"]').forEach(link => {
  link.addEventListener('click', () => {
    if (typeof gtag !== 'undefined') {
      gtag('event', 'phone_call_click', { event_category: 'CTA', event_label: link.textContent.trim() });
    }
  });
});

// ===== MULTI-STEP LEAD FORM =====
const dots  = document.querySelectorAll('.form-step-dot');
const lines = document.querySelectorAll('.form-step-line');
const bar   = document.getElementById('progressBar');
const steps = { 1: 33, 2: 66, 3: 100 };

function goToStep(n) {
  document.querySelectorAll('.form-step').forEach(s => s.classList.add('hidden'));
  document.getElementById('step' + n).classList.remove('hidden');

  dots.forEach((d, i) => {
    d.classList.toggle('active', i === n - 1);
    d.classList.toggle('done',   i <  n - 1);
  });
  lines.forEach((l, i) => l.classList.toggle('done', i < n - 1));

  if (steps[n]) bar.style.width = steps[n] + '%';
}

// ---- Validation helpers ----
function showError(id, msg) {
  const el = document.getElementById(id);
  if (el) el.textContent = msg;
}
function clearError(id) { showError(id, ''); }

function validateZip(val) {
  return /^\d{5}$/.test(val.trim());
}
function validateName(val) {
  return val.trim().length >= 2;
}
function validatePhone(val) {
  return val.replace(/\D/g, '').length >= 10;
}

// ---- Phone auto-format ----
const phoneInput = document.getElementById('phoneInput');
phoneInput.addEventListener('input', () => {
  let digits = phoneInput.value.replace(/\D/g, '').slice(0, 10);
  if (digits.length > 6)      phoneInput.value = '(' + digits.slice(0,3) + ') ' + digits.slice(3,6) + '-' + digits.slice(6);
  else if (digits.length > 3) phoneInput.value = '(' + digits.slice(0,3) + ') ' + digits.slice(3);
  else if (digits.length > 0) phoneInput.value = '(' + digits;
});

// ---- Step 1: ZIP ----
document.getElementById('zipNext').addEventListener('click', () => {
  const zip = document.getElementById('zipInput').value;
  if (!validateZip(zip)) {
    document.getElementById('zipInput').classList.add('invalid');
    showError('zipError', 'Please enter a valid 5-digit ZIP code.');
    return;
  }
  document.getElementById('zipInput').classList.remove('invalid');
  clearError('zipError');
  goToStep(2);
  document.getElementById('firstInput').focus();
});

document.getElementById('zipInput').addEventListener('keydown', e => {
  if (e.key === 'Enter') document.getElementById('zipNext').click();
});

// ---- Step 2: Name ----
document.getElementById('nameNext').addEventListener('click', () => {
  const first = document.getElementById('firstInput').value;
  const last  = document.getElementById('lastInput').value;
  let ok = true;

  if (!validateName(first)) {
    document.getElementById('firstInput').classList.add('invalid');
    showError('firstError', 'Required.');
    ok = false;
  } else {
    document.getElementById('firstInput').classList.remove('invalid');
    clearError('firstError');
  }

  if (!validateName(last)) {
    document.getElementById('lastInput').classList.add('invalid');
    showError('lastError', 'Required.');
    ok = false;
  } else {
    document.getElementById('lastInput').classList.remove('invalid');
    clearError('lastError');
  }

  if (!ok) return;
  goToStep(3);
  document.getElementById('phoneInput').focus();
});

document.getElementById('nameBack').addEventListener('click', () => goToStep(1));

// ---- Step 3: Phone + submit ----
document.getElementById('formSubmit').addEventListener('click', async () => {
  const phone = document.getElementById('phoneInput').value;
  const tcpa  = document.getElementById('tcpaCheck').checked;
  let ok = true;

  if (!validatePhone(phone)) {
    document.getElementById('phoneInput').classList.add('invalid');
    showError('phoneError', 'Please enter a valid 10-digit phone number.');
    ok = false;
  } else {
    document.getElementById('phoneInput').classList.remove('invalid');
    clearError('phoneError');
  }

  if (!tcpa) {
    showError('tcpaError', 'Please accept the consent to continue.');
    ok = false;
  } else {
    clearError('tcpaError');
  }

  if (!ok) return;

  const submitBtn = document.getElementById('formSubmit');
  submitBtn.disabled = true;
  submitBtn.textContent = 'Submitting…';

  const payload = {
    zip:   document.getElementById('zipInput').value.trim(),
    first: document.getElementById('firstInput').value.trim(),
    last:  document.getElementById('lastInput').value.trim(),
    phone: phone,
    source: window.location.hostname,
  };

  try {
    const res = await fetch('/submit-lead', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    if (!res.ok) throw new Error(`Server error ${res.status}`);
  } catch (err) {
    console.error('Lead submission error:', err);
    // Still show success UX — don't block the user over a network hiccup
  }

  // Track conversion
  if (typeof gtag !== 'undefined') {
    gtag('event', 'lead_form_submit', { event_category: 'Lead', event_label: payload.zip });
  }
  // StackAdapt conversion event
  if (typeof saq !== 'undefined') {
    saq('conv', 'ay8PT5S3JLDKAjCC565w8A');
  }

  // Show success
  bar.style.width = '100%';
  dots.forEach(d => { d.classList.remove('active'); d.classList.add('done'); });
  lines.forEach(l => l.classList.add('done'));
  document.querySelectorAll('.form-step').forEach(s => s.classList.add('hidden'));
  document.getElementById('successPhone').textContent = phone;
  document.getElementById('stepSuccess').classList.remove('hidden');
  launchConfetti();
});

document.getElementById('phoneBack').addEventListener('click', () => goToStep(2));
