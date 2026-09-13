let analysisData = null;
let selectedFile = null;

// ── Drag & Drop ──────────────────────────────────────────────────────────────
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');

uploadArea.addEventListener('click', () => fileInput.click());

uploadArea.addEventListener('dragover', e => {
  e.preventDefault();
  uploadArea.classList.add('dragging');
});

uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('dragging'));

uploadArea.addEventListener('drop', e => {
  e.preventDefault();
  uploadArea.classList.remove('dragging');
  const file = e.dataTransfer.files[0];
  if (file) handleFileSelect(file);
});

fileInput.addEventListener('change', e => {
  if (e.target.files[0]) handleFileSelect(e.target.files[0]);
});

function handleFileSelect(file) {
  const allowed = ['application/pdf', 'text/plain'];
  if (!allowed.includes(file.type) && !file.name.endsWith('.txt') && !file.name.endsWith('.pdf')) {
    alert('Please upload a PDF or TXT file.');
    return;
  }
  selectedFile = file;
  uploadArea.classList.add('file-selected');
  uploadArea.innerHTML = `
    <div class="upload-icon">✅</div>
    <p class="upload-text" style="color:#065f46;font-weight:600;">${file.name}</p>
    <p class="file-name">${(file.size / 1024).toFixed(1)} KB</p>
    <p class="upload-hint" style="margin-top:8px;">
      <label for="fileInput" class="upload-link">Change file</label>
    </p>
    <input type="file" id="fileInput" accept=".pdf,.txt" hidden />
  `;
  document.getElementById('fileInput').addEventListener('change', e => {
    if (e.target.files[0]) handleFileSelect(e.target.files[0]);
  });
}

// ── Analysis ─────────────────────────────────────────────────────────────────
async function startAnalysis() {
  const resumeText = document.getElementById('resumeText').value.trim();
  const targetRole = document.getElementById('targetRole').value.trim();

  if (!selectedFile && !resumeText) {
    alert('Please upload a resume file or paste your resume text.');
    return;
  }

  document.getElementById('uploadSection').classList.add('hidden');
  document.getElementById('loadingSection').classList.remove('hidden');

  const btn = document.getElementById('analyzeBtn');
  btn.disabled = true;

  // Animate steps
  const stepTimings = [0, 2000, 5000, 9000, 14000];
  const steps = document.querySelectorAll('.step');

  stepTimings.forEach((delay, i) => {
    setTimeout(() => {
      steps.forEach((s, j) => {
        s.classList.toggle('active', j === i);
        if (j < i) s.classList.add('done');
      });
    }, delay);
  });

  try {
    const formData = new FormData();
    if (selectedFile) {
      formData.append('resume', selectedFile);
    } else {
      formData.append('resumeText', resumeText);
    }
    if (targetRole) formData.append('targetRole', targetRole);

    const response = await fetch('/api/analyze', {
      method: 'POST',
      body: formData
    });

    if (!response.ok) {
      const err = await response.json();
      throw new Error(err.error || 'Analysis failed');
    }

    // Stream SSE
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    let jsonText = '';

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;

      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split('\n');
      buffer = lines.pop();

      for (const line of lines) {
        if (!line.startsWith('data: ')) continue;
        const payload = JSON.parse(line.slice(6));
        if (payload.type === 'delta') jsonText += payload.text;
        if (payload.type === 'error') throw new Error(payload.message);
        if (payload.type === 'done') break;
      }
    }

    const data = parseJSON(jsonText);
    if (!data) throw new Error('Invalid response from AI. Please try again.');

    analysisData = data;
    renderResults(data);

    document.getElementById('loadingSection').classList.add('hidden');
    document.getElementById('resultsSection').classList.remove('hidden');
    window.scrollTo({ top: 0, behavior: 'smooth' });

  } catch (err) {
    document.getElementById('loadingSection').classList.add('hidden');
    document.getElementById('uploadSection').classList.remove('hidden');
    btn.disabled = false;
    alert('Error: ' + err.message);
  }
}

function parseJSON(text) {
  // Strip markdown code fences if present
  const cleaned = text.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/i, '').trim();
  try {
    return JSON.parse(cleaned);
  } catch {
    // Try to find JSON object
    const start = cleaned.indexOf('{');
    const end = cleaned.lastIndexOf('}');
    if (start !== -1 && end !== -1) {
      try { return JSON.parse(cleaned.slice(start, end + 1)); } catch {}
    }
    return null;
  }
}

// ── Render ────────────────────────────────────────────────────────────────────
function renderResults(d) {
  renderScoreHero(d.overall);
  renderScoreBars(d.scores);
  renderIssuesTab(d);
  renderSectionsTab(d.sections);
  renderRewritesTab(d);
  renderLevelUpTab(d);
}

function renderScoreHero(overall) {
  document.getElementById('overallScore').textContent = overall.score;
  document.getElementById('scoreGrade').textContent = overall.grade;
  document.getElementById('scoreHeadline').textContent = overall.headline;
  document.getElementById('atsScore').textContent = overall.ats_score;
  document.getElementById('hireLikelihood').textContent = overall.hiring_likelihood;
  document.getElementById('recruiterTime').textContent = overall.time_to_review + 's';

  const circumference = 2 * Math.PI * 80;
  const offset = circumference - (overall.score / 100) * circumference;
  setTimeout(() => {
    document.getElementById('ringFill').style.strokeDashoffset = offset;
  }, 200);

  const scoreEl = document.getElementById('overallScore');
  const score = overall.score;
  if (score >= 80) scoreEl.style.color = '#34d399';
  else if (score >= 60) scoreEl.style.color = '#fbbf24';
  else scoreEl.style.color = '#f87171';
}

function renderScoreBars(scores) {
  const labels = {
    impact: 'Impact',
    clarity: 'Clarity',
    relevance: 'Relevance',
    formatting: 'Formatting',
    keywords: 'Keywords',
    achievements: 'Achievements'
  };

  const container = document.getElementById('scoreBars');
  container.innerHTML = Object.entries(scores).map(([key, val]) => {
    const cls = val >= 75 ? 'high' : val >= 50 ? 'mid' : 'low';
    return `
      <div class="score-bar-item">
        <span class="bar-label">${labels[key] || key}</span>
        <div class="bar-track">
          <div class="bar-fill ${cls}" data-width="${val}" style="width:0%"></div>
        </div>
        <span class="bar-score">${val}</span>
      </div>`;
  }).join('');

  setTimeout(() => {
    container.querySelectorAll('.bar-fill').forEach(el => {
      el.style.width = el.dataset.width + '%';
    });
  }, 100);
}

function renderIssuesTab(d) {
  const strengthsList = document.getElementById('strengthsList');
  strengthsList.innerHTML = (d.strengths || []).map(s => `<li>${s}</li>`).join('');

  const quickList = document.getElementById('quickWinsList');
  quickList.innerHTML = (d.quick_wins || []).map(w => `<li>${w}</li>`).join('');

  const issuesList = document.getElementById('issuesList');
  issuesList.innerHTML = (d.critical_issues || []).map(issue => `
    <div class="issue-item ${issue.severity}">
      <div class="issue-header">
        <span class="severity-badge ${issue.severity}">${issue.severity}</span>
        <span class="issue-title">${issue.issue}</span>
      </div>
      <p class="issue-detail">${issue.detail}</p>
      <div class="issue-fix">
        <span class="fix-label">Fix:</span>${issue.fix}
      </div>
    </div>`).join('');
}

function renderSectionsTab(sections) {
  const icons = {
    contact: '📇',
    summary: '📝',
    experience: '💼',
    education: '🎓',
    skills: '🛠️'
  };

  const container = document.getElementById('sectionsContent');
  container.innerHTML = Object.entries(sections).map(([key, sec]) => {
    const scoreClass = sec.score >= 75 ? 'high' : sec.score >= 50 ? 'mid' : 'low';
    let extra = '';

    if (sec.missing && sec.missing.length > 0) {
      extra += `<div class="missing-items">${sec.missing.map(m => `<span class="missing-tag">${m}</span>`).join('')}</div>`;
    }

    if (key === 'skills') {
      if (sec.missing_keywords && sec.missing_keywords.length > 0) {
        extra += `<p style="font-size:13px;font-weight:600;margin-top:12px;margin-bottom:6px;color:#6366f1;">Missing Keywords:</p>
          <div class="keyword-tags">${sec.missing_keywords.slice(0, 8).map(k => `<span class="keyword-tag">${k}</span>`).join('')}</div>`;
      }
    }

    return `
      <div class="section-item">
        <div class="section-header">
          <span class="section-name">${icons[key] || '📄'} ${key.charAt(0).toUpperCase() + key.slice(1)}</span>
          <span class="section-score-badge ${scoreClass}">${sec.score}</span>
        </div>
        <p class="section-feedback">${sec.feedback}</p>
        ${extra}
      </div>`;
  }).join('');
}

function renderRewritesTab(d) {
  // Summary rewrite
  const summaryEl = document.getElementById('summaryRewrite');
  const summaryRewrite = d.sections?.summary?.rewrite;
  if (summaryRewrite) {
    summaryEl.innerHTML = `
      <div class="summary-rewrite-card">
        <h3 class="card-title">✍️ Rewritten Professional Summary</h3>
        <p class="card-desc" style="margin-bottom:16px;">AI-crafted summary to make a stronger first impression</p>
        <p>${summaryRewrite}</p>
      </div>`;
  }

  // Bullet rewrites
  const bullets = d.sections?.experience?.bullet_rewrites || [];
  const bulletContainer = document.getElementById('bulletRewrites');
  if (bullets.length === 0) {
    bulletContainer.innerHTML = '<p style="color:var(--text-muted);font-size:14px;">No specific bullet rewrites generated. Your bullets may already be strong, or the experience section may need more detail.</p>';
  } else {
    bulletContainer.innerHTML = bullets.map(b => `
      <div class="bullet-rewrite">
        <div class="rewrite-before">
          <span class="rewrite-label">Before</span>
          ${b.original}
        </div>
        <div class="rewrite-after">
          <span class="rewrite-label">After</span>
          ${b.improved}
        </div>
        <div class="rewrite-why">💡 ${b.why}</div>
      </div>`).join('');
  }

  // Keywords
  const keywords = d.sections?.skills?.missing_keywords || [];
  const keywordContainer = document.getElementById('keywordTags');
  keywordContainer.innerHTML = keywords.length > 0
    ? keywords.map(k => `<span class="keyword-tag">${k}</span>`).join('')
    : '<p style="color:var(--text-muted);font-size:14px;">Your keywords look solid for the target role.</p>';
}

function renderLevelUpTab(d) {
  const tips = d.level_up_tips || [];
  const container = document.getElementById('levelUpList');
  container.innerHTML = tips.map(tip => `
    <div class="level-up-item">
      <div class="level-up-header">
        <span class="level-up-tip">${tip.tip}</span>
        <div class="level-up-badges">
          <span class="impact-badge ${tip.impact}">${tip.impact} Impact</span>
          <span class="effort-badge ${tip.effort}">${tip.effort}</span>
        </div>
      </div>
      <div class="level-up-example">${tip.example}</div>
    </div>`).join('');

  const tailoring = d.tailoring_suggestions || [];
  if (tailoring.length > 0) {
    document.getElementById('tailoringCard').classList.remove('hidden');
    document.getElementById('tailoringList').innerHTML = tailoring.map(t =>
      `<div class="tailoring-item">${t}</div>`).join('');
  }
}

// ── Tabs ──────────────────────────────────────────────────────────────────────
function showTab(name) {
  document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
  document.querySelectorAll('.tab').forEach(el => el.classList.remove('active'));
  document.getElementById('tab-' + name).classList.remove('hidden');
  event.target.classList.add('active');
}

// ── Reset ─────────────────────────────────────────────────────────────────────
function resetApp() {
  selectedFile = null;
  analysisData = null;

  document.getElementById('resumeText').value = '';
  document.getElementById('targetRole').value = '';
  document.getElementById('resultsSection').classList.add('hidden');
  document.getElementById('uploadSection').classList.remove('hidden');
  document.getElementById('analyzeBtn').disabled = false;

  const ua = document.getElementById('uploadArea');
  ua.classList.remove('file-selected', 'dragging');
  ua.innerHTML = `
    <div class="upload-icon">📄</div>
    <p class="upload-text">Drop your resume here or <label for="fileInput" class="upload-link">browse files</label></p>
    <p class="upload-hint">Supports PDF and TXT · Max 10MB</p>
    <input type="file" id="fileInput" accept=".pdf,.txt" hidden />
  `;
  document.getElementById('fileInput').addEventListener('change', e => {
    if (e.target.files[0]) handleFileSelect(e.target.files[0]);
  });

  // Reset tab
  document.querySelectorAll('.tab').forEach((t, i) => t.classList.toggle('active', i === 0));
  document.querySelectorAll('.tab-content').forEach((t, i) => t.classList.toggle('hidden', i !== 0));

  window.scrollTo({ top: 0, behavior: 'smooth' });
}
