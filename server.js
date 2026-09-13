require('dotenv').config();
const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const Anthropic = require('@anthropic-ai/sdk');

const app = express();
const PORT = process.env.PORT || 3000;
const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const allowed = ['application/pdf', 'text/plain'];
    if (allowed.includes(file.mimetype)) cb(null, true);
    else cb(new Error('Only PDF and TXT files allowed'));
  }
});

app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

async function extractTextFromPDF(buffer) {
  const pdfParse = require('pdf-parse');
  const data = await pdfParse(buffer);
  return data.text;
}

const SYSTEM_PROMPT = `You are an elite professional resume consultant and recruiter with 15+ years of experience at top-tier companies including Fortune 500 corporations, leading tech firms, and prestigious consulting houses. You have reviewed over 10,000 resumes and have deep expertise in:

- ATS (Applicant Tracking System) optimization
- Industry-specific resume standards across tech, finance, healthcare, marketing, and more
- Executive and entry-level positioning
- Quantifiable achievement framing
- Keyword optimization for modern job markets
- LinkedIn profile alignment

Your analysis is direct, actionable, and highly detailed. You identify both fatal flaws and subtle improvements that transform good resumes into exceptional ones.

ALWAYS respond with valid JSON only — no markdown, no prose outside JSON. Structure your response exactly as specified.`;

function buildAnalysisPrompt(resumeText, targetRole) {
  const roleContext = targetRole ? `The candidate is targeting: ${targetRole}` : 'Analyze for general professional positioning.';

  return `${roleContext}

Analyze this resume comprehensively and return ONLY valid JSON with this exact structure:

{
  "overall": {
    "score": <0-100 integer>,
    "grade": "<A+/A/A-/B+/B/B-/C+/C/C-/D/F>",
    "headline": "<one punchy sentence summary of the resume's current state>",
    "ats_score": <0-100 integer>,
    "hiring_likelihood": "<High/Medium/Low>",
    "time_to_review": "<estimated seconds a recruiter would spend: 3-45>"
  },
  "scores": {
    "impact": <0-100>,
    "clarity": <0-100>,
    "relevance": <0-100>,
    "formatting": <0-100>,
    "keywords": <0-100>,
    "achievements": <0-100>
  },
  "strengths": [
    "<specific strength with why it works>",
    "<specific strength>",
    "<specific strength>"
  ],
  "critical_issues": [
    {
      "severity": "<critical|warning|suggestion>",
      "issue": "<concise issue title>",
      "detail": "<why this hurts the candidate>",
      "fix": "<specific actionable fix>"
    }
  ],
  "sections": {
    "contact": {
      "score": <0-100>,
      "status": "<good|needs_work|missing>",
      "feedback": "<specific feedback>",
      "missing": ["<any missing contact elements>"]
    },
    "summary": {
      "score": <0-100>,
      "status": "<good|needs_work|missing>",
      "feedback": "<specific feedback>",
      "rewrite": "<full rewritten professional summary if score < 80, else null>"
    },
    "experience": {
      "score": <0-100>,
      "status": "<good|needs_work|missing>",
      "feedback": "<overall experience section feedback>",
      "bullet_rewrites": [
        {
          "original": "<exact original bullet>",
          "improved": "<rewritten bullet with metrics and impact>",
          "why": "<brief explanation of improvement>"
        }
      ]
    },
    "education": {
      "score": <0-100>,
      "status": "<good|needs_work|missing>",
      "feedback": "<specific feedback>"
    },
    "skills": {
      "score": <0-100>,
      "status": "<good|needs_work|missing>",
      "feedback": "<specific feedback>",
      "missing_keywords": ["<important keywords missing for the role>"],
      "remove": ["<weak or outdated skills to remove>"]
    }
  },
  "level_up_tips": [
    {
      "tip": "<specific actionable tip>",
      "impact": "<High|Medium|Low>",
      "effort": "<Easy|Medium|Hard>",
      "example": "<concrete example of how to implement>"
    }
  ],
  "quick_wins": [
    "<change that takes under 5 minutes and immediately improves the resume>"
  ],
  "tailoring_suggestions": [
    "<specific suggestion tailored to the target role if provided>"
  ]
}

RESUME TEXT:
${resumeText}`;
}

app.post('/api/analyze', upload.single('resume'), async (req, res) => {
  try {
    let resumeText = '';

    if (req.file) {
      if (req.file.mimetype === 'application/pdf') {
        resumeText = await extractTextFromPDF(req.file.buffer);
      } else {
        resumeText = req.file.buffer.toString('utf-8');
      }
    } else if (req.body.resumeText) {
      resumeText = req.body.resumeText;
    } else {
      return res.status(400).json({ error: 'No resume provided' });
    }

    if (!resumeText.trim() || resumeText.trim().length < 50) {
      return res.status(400).json({ error: 'Resume text is too short or empty' });
    }

    const targetRole = req.body.targetRole || '';

    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');

    const stream = await client.messages.stream({
      model: 'claude-opus-4-8',
      max_tokens: 8000,
      thinking: { type: 'adaptive' },
      system: SYSTEM_PROMPT,
      messages: [
        {
          role: 'user',
          content: buildAnalysisPrompt(resumeText, targetRole)
        }
      ]
    });

    let fullText = '';

    for await (const event of stream) {
      if (event.type === 'content_block_delta' && event.delta.type === 'text_delta') {
        fullText += event.delta.text;
        res.write(`data: ${JSON.stringify({ type: 'delta', text: event.delta.text })}\n\n`);
      }
    }

    res.write(`data: ${JSON.stringify({ type: 'done' })}\n\n`);
    res.end();
  } catch (err) {
    console.error('Analysis error:', err);
    if (!res.headersSent) {
      res.status(500).json({ error: err.message || 'Analysis failed' });
    } else {
      res.write(`data: ${JSON.stringify({ type: 'error', message: err.message })}\n\n`);
      res.end();
    }
  }
});

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
  console.log(`Resume Analyzer running at http://localhost:${PORT}`);
});
