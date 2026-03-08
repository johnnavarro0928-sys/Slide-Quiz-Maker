import { useState, useCallback, useEffect, useRef } from "react";
import * as mammoth from "mammoth";

// ── Slide color palette (eye-friendly, high contrast) ──────────
const Q_BG      = "1C2B3A";   // soft navy - question slide bg
const Q_ACCENT  = "4EA8DE";   // sky blue - question accents
const Q_NUM_BG  = "F0A500";   // amber - number badge
const Q_HEADER  = "A8C8E8";   // light blue - header label
const Q_TEXT    = "F0F4F8";   // near-white - question text
const Q_LINE    = "2E4A63";   // divider
const A_BG      = "16293B";   // answer slide bg
const A_ACCENT  = "3DD68C";   // mint green - correct answer
const A_BOX     = "1E3D2F";   // answer highlight box
const A_RAT_BG  = "132030";   // rationale box bg
const A_RAT_LN  = "2A5C45";   // rationale box border
const A_TEXT    = "E8F5EE";   // answer text
const A_MUTED   = "8ECFAA";   // muted green label
const WHITE     = "FFFFFF";
const AMBER     = "F0A500";
const CHOICE_COLORS = ["1E4D8C", "1C4D3B", "4D2060", "4D3800"];

export default function App() {
  const [stage, setStage] = useState("upload");
  const [file, setFile] = useState(null);
  const [fileData, setFileData] = useState(null);
  const [questions, setQuestions] = useState([]);
  const [testTitle, setTestTitle] = useState("Test");
  const [randomize, setRandomize] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const [status, setStatus] = useState("");
  const [error, setError] = useState("");
  const [expandedQ, setExpandedQ] = useState(null);
  const fileInputRef = useRef(null);
  const pptxReady = useRef(false);
  const pdfjsReady = useRef(null);

  useEffect(() => {
    const script = document.createElement("script");
    script.src = "https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js";
    script.onload = () => { pptxReady.current = true; };
    document.head.appendChild(script);
  }, []);

  const fileDataRef = useRef(null);

  const loadPdfJs = useCallback(async () => {
    if (window.pdfjsLib) return window.pdfjsLib;
    if (pdfjsReady.current) return pdfjsReady.current;

    pdfjsReady.current = new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
      script.async = true;
      script.onload = () => {
        if (!window.pdfjsLib) {
          reject(new Error("PDF parser failed to initialize."));
          return;
        }
        window.pdfjsLib.GlobalWorkerOptions.workerSrc =
          "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
        resolve(window.pdfjsLib);
      };
      script.onerror = () => reject(new Error("Failed to load PDF parser."));
      document.head.appendChild(script);
    });

    return pdfjsReady.current;
  }, []);

  const extractTextFromPdf = useCallback(async (selectedFile) => {
    const pdfjsLib = await loadPdfJs();
    const arrayBuffer = await selectedFile.arrayBuffer();
    const pdfDoc = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const pageText = [];

    for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum += 1) {
      const page = await pdfDoc.getPage(pageNum);
      const textContent = await page.getTextContent();
      pageText.push(textContent.items.map((item) => item.str || "").join(" "));
    }

    return pageText.join("\n").replace(/[ \t]+/g, " ").trim();
  }, [loadPdfJs]);

  const normalizeAnswer = useCallback((rawAnswer, choices) => {
    const answer = (rawAnswer || "").trim();
    if (!answer) return "Not provided";

    const letterMatch = answer.match(/^\(?([A-H])\)?$/i);
    if (letterMatch && Array.isArray(choices) && choices.length > 0) {
      const target = letterMatch[1].toUpperCase();
      const found = choices.find((choice) => choice.toUpperCase().startsWith(`${target}.`));
      return found || `${target}.`;
    }

    return answer;
  }, []);

  const parseQuestionsFromText = useCallback((rawText) => {
    const text = (rawText || "").replace(/\r/g, "");
    const lines = text.split("\n").map((line) => line.trimEnd());
    const questionRe = /^\s*(?:q(?:uestion)?\s*)?(\d{1,4})\s*[\).:\-]\s*(.+)$/i;
    const choiceRe = /^\s*([A-H])\s*[\).:\-]\s*(.+)$/i;
    const answerRe = /^\s*(?:answer|correct answer|ans(?:wer)?(?:\s*key)?)\s*[:\-]\s*(.+)$/i;
    const rationaleRe = /^\s*(?:rationale|explanation|reason(?:ing)?)\s*[:\-]\s*(.*)$/i;

    const firstQuestionIdx = lines.findIndex((line) => questionRe.test(line));
    let title = "Test";
    if (firstQuestionIdx > 0) {
      const maybeTitle = lines.slice(0, firstQuestionIdx).find((line) => line.trim().length > 0);
      if (maybeTitle) {
        title = maybeTitle.replace(/^(?:title|quiz|test)\s*[:\-]\s*/i, "").trim() || "Test";
      }
    }

    const extracted = [];
    let current = null;
    let section = "question";

    const finalizeCurrent = () => {
      if (!current || !current.question.trim()) return;
      extracted.push({
        number: current.number || extracted.length + 1,
        question: current.question.replace(/\s+/g, " ").trim(),
        choices: current.choices.map((choice) => choice.replace(/\s+/g, " ").trim()),
        answer: normalizeAnswer(current.answer, current.choices),
        rationale: current.rationale.replace(/\s+/g, " ").trim() || "No rationale provided."
      });
      current = null;
      section = "question";
    };

    lines.forEach((rawLine) => {
      const line = rawLine.trim();
      if (!line) return;

      const qMatch = line.match(questionRe);
      if (qMatch) {
        finalizeCurrent();
        current = {
          number: Number(qMatch[1]),
          question: qMatch[2].trim(),
          choices: [],
          answer: "",
          rationale: ""
        };
        section = "question";
        return;
      }

      if (!current) return;

      const choiceMatch = line.match(choiceRe);
      if (choiceMatch) {
        current.choices.push(`${choiceMatch[1].toUpperCase()}. ${choiceMatch[2].trim()}`);
        section = "choice";
        return;
      }

      const answerMatch = line.match(answerRe);
      if (answerMatch) {
        current.answer = answerMatch[1].trim();
        section = "answer";
        return;
      }

      const rationaleMatch = line.match(rationaleRe);
      if (rationaleMatch) {
        current.rationale = rationaleMatch[1].trim();
        section = "rationale";
        return;
      }

      if (section === "choice" && current.choices.length > 0) {
        const lastIdx = current.choices.length - 1;
        current.choices[lastIdx] = `${current.choices[lastIdx]} ${line}`.trim();
        return;
      }

      if (section === "answer") {
        current.answer = `${current.answer} ${line}`.trim();
        return;
      }

      if (section === "rationale") {
        current.rationale = `${current.rationale} ${line}`.trim();
        return;
      }

      current.question = `${current.question} ${line}`.trim();
    });

    finalizeCurrent();

    if (extracted.length === 0) {
      throw new Error(
        "No questions were detected. Please use a format like '1. ...', choices 'A. ...', and include 'Answer:' / 'Rationale:'."
      );
    }

    return { title, questions: extracted };
  }, [normalizeAnswer]);

  const readFile = useCallback(async (selectedFile) => {
    if (!selectedFile) return;
    setFile(selectedFile);
    setError("");
    setFileData(null);
    fileDataRef.current = null;
    setStatus("Reading file...");

    try {
      const name = selectedFile.name.toLowerCase();
      const isPdf = selectedFile.type === "application/pdf";
      const isDocx = name.endsWith(".docx") || name.endsWith(".doc");

      if (isPdf) {
        const text = await extractTextFromPdf(selectedFile);
        if (!text || text.trim().length < 10) throw new Error("Could not extract text from this PDF file.");
        fileDataRef.current = text;
        setFileData(text);
      } else if (isDocx) {
        const arrayBuffer = await selectedFile.arrayBuffer();
        const result = await mammoth.extractRawText({ arrayBuffer });
        const text = result.value;
        if (!text || text.trim().length < 10) throw new Error("Could not extract text from this Word file.");
        fileDataRef.current = text;
        setFileData(text);
      } else {
        // Plain text / csv / etc
        const reader = new FileReader();
        await new Promise((resolve, reject) => {
          reader.onload = (e) => {
            fileDataRef.current = e.target.result;
            setFileData(e.target.result);
            resolve();
          };
          reader.onerror = reject;
          reader.readAsText(selectedFile);
        });
      }
      setStatus("✅ File ready!");
      setTimeout(() => setStatus(""), 2000);
    } catch (err) {
      setError("Failed to read file: " + err.message);
      setStatus("");
    }
  }, [extractTextFromPdf]);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    setDragOver(false);
    const f = e.dataTransfer.files[0];
    if (f) readFile(f);
  }, [readFile]);

  const [debugLog, setDebugLog] = useState([]);
  const log = (msg) => setDebugLog(prev => [...prev, msg]);

  const extractQuestions = async () => {
    const currentData = fileDataRef.current;

    if (!currentData) {
      setError("File not ready yet — please wait a moment and try again.");
      return;
    }

    setStage("processing");
    setStatus("Extracting questions from file...");
    setError("");
    setDebugLog([]);

    try {
      log(`Source text length: ${currentData.length}`);
      const parsed = parseQuestionsFromText(currentData);
      const extractedQs = parsed.questions || [];
      log(`Parsed questions: ${extractedQs.length}`);

      if (extractedQs.length === 0) {
        throw new Error("No valid questions found in the uploaded test paper.");
      }

      setTestTitle(parsed.title || "Test");
      setQuestions(extractedQs);
      setStage("preview");
    } catch (err) {
      setError(err.message || String(err));
      setStage("upload");
    }
  };

  const generatePPTX = async () => {
    if (!pptxReady.current || !window.PptxGenJS) {
      setError("Slide engine still loading, please try again in a second.");
      return;
    }
    setStage("generating");
    setStatus("Building your presentation...");

    const finalQs = randomize ? [...questions].sort(() => Math.random() - 0.5) : questions;

    try {
      const pptx = new window.PptxGenJS();
      pptx.layout = "LAYOUT_16x9";
      pptx.title = testTitle;

      // ── TITLE SLIDE ─────────────────────────────────────────────
      const ts = pptx.addSlide();
      ts.background = { color: "1A2740" };
      // top & bottom amber bars
      ts.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0,     w: 10, h: 0.18, fill: { color: AMBER } });
      ts.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 5.445, w: 10, h: 0.18, fill: { color: AMBER } });
      // decorative circle
      ts.addShape(pptx.shapes.OVAL, { x: 6.8, y: 0.3, w: 4.5, h: 4.5, fill: { color: "253A5E", transparency: 60 } });
      ts.addShape(pptx.shapes.OVAL, { x: -1.5, y: 3.0,  w: 3.5, h: 3.5, fill: { color: "253A5E", transparency: 65 } });
      // label + title
      ts.addText("QUIZ PRESENTATION", {
        x: 0.6, y: 1.4, w: 8.8, h: 0.5,
        fontSize: 14, color: AMBER, bold: true, charSpacing: 6, align: "center", fontFace: "Calibri"
      });
      ts.addText(testTitle, {
        x: 0.6, y: 2.0, w: 8.8, h: 1.8,
        fontSize: 44, color: WHITE, bold: true, align: "center", fontFace: "Cambria", lineSpacingMultiple: 1.1
      });
      ts.addText(`${finalQs.length} Questions  •  ${randomize ? "Randomized Order" : "Original Order"}`, {
        x: 0.6, y: 4.1, w: 8.8, h: 0.5,
        fontSize: 18, color: Q_HEADER, align: "center", fontFace: "Calibri"
      });

      // ── HELPERS ────────────────────────────────────────────────
      // Estimate how many lines a string needs given a text box width and font size.
      // charsPerInch is approximate for Calibri: ~1.6 chars per pt per inch.
      const estLines = (text, boxW, fontSize) => {
        const charsPerLine = Math.floor((boxW * 72) / (fontSize * 0.55));
        return Math.max(1, Math.ceil(text.length / charsPerLine));
      };
      // Convert lines to inches given fontSize and lineSpacing multiplier
      const linesToIn = (lines, fontSize, lsm) => (lines * fontSize * lsm) / 72;

      // Pick the largest font size that keeps text within maxH inches
      const fitFontSize = (text, boxW, maxH, sizes, lsm) => {
        for (const sz of sizes) {
          const lines = estLines(text, boxW, sz);
          if (linesToIn(lines, sz, lsm) <= maxH) return sz;
        }
        return sizes[sizes.length - 1]; // smallest
      };

      // ── QUESTION + ANSWER SLIDES ──────────────────────────────
      finalQs.forEach((q, idx) => {
        const qNum = idx + 1;
        const hasChoices = Array.isArray(q.choices) && q.choices.length > 0;
        const SLIDE_H = 5.625;
        const HEADER_H = 1.05;          // fixed top bar height
        const MARGIN = 0.35;
        const CONTENT_W = 9.3;          // usable width
        const CONTENT_X = 0.35;
        const CONTENT_TOP = HEADER_H + 0.1;  // y where content starts
        const AVAIL_H = SLIDE_H - CONTENT_TOP - 0.1; // total usable height below header

        // ── QUESTION SLIDE ──────────────────────────────────────
        const qs = pptx.addSlide();
        qs.background = { color: Q_BG };

        // header bar
        qs.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: HEADER_H, fill: { color: "162030" } });
        qs.addShape(pptx.shapes.RECTANGLE, { x: 0, y: HEADER_H, w: 10, h: 0.05, fill: { color: Q_ACCENT } });

        // number badge
        qs.addShape(pptx.shapes.OVAL, { x: 0.28, y: 0.1, w: 0.85, h: 0.85, fill: { color: Q_NUM_BG } });
        qs.addText(String(qNum), {
          x: 0.28, y: 0.1, w: 0.85, h: 0.85,
          fontSize: 28, color: "111820", bold: true,
          align: "center", valign: "middle", fontFace: "Cambria", margin: 0
        });
        qs.addText("QUESTION", {
          x: 1.28, y: 0.27, w: 4, h: 0.5,
          fontSize: 20, color: Q_ACCENT, bold: true, charSpacing: 5, fontFace: "Calibri", valign: "middle"
        });

        if (!hasChoices) {
          // ── No choices: question fills entire content area ──
          const qFontSize = fitFontSize(q.question, CONTENT_W, AVAIL_H, [40,36,32,28,24,22,20], 1.3);
          qs.addText(q.question, {
            x: CONTENT_X, y: CONTENT_TOP, w: CONTENT_W, h: AVAIL_H,
            fontSize: qFontSize, color: Q_TEXT, fontFace: "Calibri",
            valign: "middle", wrap: true, lineSpacingMultiple: 1.3
          });
        } else {
          // ── With choices: split space between question and choice grid ──
          const cols = q.choices.length > 2 ? 2 : 1;
          const rows = Math.ceil(q.choices.length / cols);
          const GAP = 0.1;
          const BADGE_W = 0.65;

          // Minimum card height to keep choice text at 22pt (1 line)
          const MIN_CARD_H = 0.85;
          const choicesBlockH = rows * MIN_CARD_H + (rows - 1) * GAP;

          // Question gets the remaining space
          const Q_GAP = 0.12; // gap between question and choices
          const qMaxH = AVAIL_H - choicesBlockH - Q_GAP;
          const qFontSize = fitFontSize(q.question, CONTENT_W, qMaxH, [32,28,26,24,22,20,18], 1.25);
          const qLines = estLines(q.question, CONTENT_W, qFontSize);
          const qActualH = Math.min(qMaxH, linesToIn(qLines, qFontSize, 1.25) + 0.15);

          qs.addText(q.question, {
            x: CONTENT_X, y: CONTENT_TOP, w: CONTENT_W, h: qActualH,
            fontSize: qFontSize, color: Q_TEXT, fontFace: "Calibri",
            valign: "middle", wrap: true, lineSpacingMultiple: 1.25
          });

          // Recalculate choice block using remaining real estate
          const choiceStartY = CONTENT_TOP + qActualH + Q_GAP;
          const choiceAvailH = SLIDE_H - choiceStartY - 0.08;
          const cardH = Math.max(MIN_CARD_H, Math.min(1.2, (choiceAvailH - (rows - 1) * GAP) / rows));
          const cardW = cols === 2 ? (CONTENT_W - GAP) / 2 : CONTENT_W;
          // Font size for choices based on card height and longest choice
          const longestChoice = q.choices.reduce((a, b) => a.length > b.length ? a : b, "");
          const choiceFontSize = fitFontSize(
            longestChoice.replace(/^[A-Da-d][.)]\s*/, ""),
            cardW - BADGE_W - 0.15, cardH - 0.06,
            [28,26,24,22,20,18], 1.15
          );

          q.choices.forEach((ch, ci) => {
            const col = Math.floor(ci / rows);
            const row = ci % rows;
            const cx = CONTENT_X + col * (cardW + GAP);
            const cy = choiceStartY + row * (cardH + GAP);
            const bg = CHOICE_COLORS[ci % 4];
            const letter = ["A","B","C","D"][ci] || String(ci + 1);

            qs.addShape(pptx.shapes.RECTANGLE, {
              x: cx, y: cy, w: cardW, h: cardH,
              fill: { color: bg, transparency: 25 }, line: { color: Q_ACCENT, width: 1 }
            });
            qs.addShape(pptx.shapes.RECTANGLE, {
              x: cx, y: cy, w: BADGE_W, h: cardH,
              fill: { color: bg, transparency: 5 }
            });
            qs.addText(letter, {
              x: cx, y: cy, w: BADGE_W, h: cardH,
              fontSize: choiceFontSize + 2, color: WHITE, bold: true,
              align: "center", valign: "middle", fontFace: "Cambria", margin: 0
            });
            qs.addText(ch.replace(/^[A-Da-d][.)]\s*/, ""), {
              x: cx + BADGE_W + 0.08, y: cy, w: cardW - BADGE_W - 0.12, h: cardH,
              fontSize: choiceFontSize, color: Q_TEXT, fontFace: "Calibri",
              valign: "middle", wrap: true, lineSpacingMultiple: 1.15
            });
          });
        }

        // ── ANSWER SLIDE ────────────────────────────────────────
        // Fixed zones: header | answer box | rationale box
        // Heights adapt to content length
        const as = pptx.addSlide();
        as.background = { color: A_BG };

        // header bar
        as.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: HEADER_H, fill: { color: "0f2018" } });
        as.addShape(pptx.shapes.RECTANGLE, { x: 0, y: HEADER_H, w: 10, h: 0.05, fill: { color: A_ACCENT } });

        as.addShape(pptx.shapes.RECTANGLE, { x: 0.28, y: 0.12, w: 3.0, h: 0.82, fill: { color: A_ACCENT } });
        as.addText("✓  ANSWER", {
          x: 0.28, y: 0.12, w: 3.0, h: 0.82,
          fontSize: 24, color: "0c1c10", bold: true,
          align: "center", valign: "middle", fontFace: "Calibri", margin: 0
        });
        as.addText(`Question ${qNum}`, {
          x: 3.45, y: 0.28, w: 3.5, h: 0.5,
          fontSize: 20, color: A_MUTED, fontFace: "Calibri", valign: "middle"
        });

        // ── answer box — fixed height 1.38", answer text auto-sized ──
        const ANS_BOX_Y = HEADER_H + 0.12;
        const ANS_BOX_H = 1.38;
        const answerText = q.answer || "See above";
        const ansFontSize = fitFontSize(answerText, CONTENT_W - 0.4, ANS_BOX_H - 0.46, [34,30,28,26,24,22], 1.2);

        as.addShape(pptx.shapes.RECTANGLE, {
          x: CONTENT_X, y: ANS_BOX_Y, w: CONTENT_W, h: ANS_BOX_H,
          fill: { color: A_BOX }, line: { color: A_ACCENT, width: 2.5 }
        });
        as.addText("CORRECT ANSWER", {
          x: CONTENT_X + 0.18, y: ANS_BOX_Y + 0.06, w: 5, h: 0.34,
          fontSize: 14, color: A_ACCENT, bold: true, charSpacing: 3, fontFace: "Calibri"
        });
        as.addText(answerText, {
          x: CONTENT_X + 0.18, y: ANS_BOX_Y + 0.38, w: CONTENT_W - 0.3, h: ANS_BOX_H - 0.44,
          fontSize: ansFontSize, color: WHITE, bold: true,
          fontFace: "Calibri", valign: "middle", wrap: true, lineSpacingMultiple: 1.2
        });

        // ── rationale — fills all remaining space ──
        const RAT_LABEL_Y = ANS_BOX_Y + ANS_BOX_H + 0.14;
        const RAT_LABEL_H = 0.38;
        const RAT_BOX_Y = RAT_LABEL_Y + RAT_LABEL_H + 0.06;
        const RAT_BOX_H = SLIDE_H - RAT_BOX_Y - 0.1;
        const rationaleText = q.rationale || "No rationale provided.";
        const ratFontSize = fitFontSize(rationaleText, CONTENT_W - 0.3, RAT_BOX_H - 0.18, [26,24,22,20,18,16], 1.4);

        as.addText("RATIONALE", {
          x: CONTENT_X, y: RAT_LABEL_Y, w: 4, h: RAT_LABEL_H,
          fontSize: 17, color: A_ACCENT, bold: true, charSpacing: 4, fontFace: "Calibri"
        });
        as.addShape(pptx.shapes.RECTANGLE, {
          x: CONTENT_X, y: RAT_BOX_Y, w: CONTENT_W, h: RAT_BOX_H,
          fill: { color: A_RAT_BG }, line: { color: A_RAT_LN, width: 1.5 }
        });
        as.addText(rationaleText, {
          x: CONTENT_X + 0.18, y: RAT_BOX_Y + 0.1, w: CONTENT_W - 0.3, h: RAT_BOX_H - 0.18,
          fontSize: ratFontSize, color: A_TEXT, fontFace: "Calibri",
          valign: "top", wrap: true, lineSpacingMultiple: 1.4
        });
      });

      // ── END SLIDE ──────────────────────────────────────────────
      const es = pptx.addSlide();
      es.background = { color: "1A2740" };
      es.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0,     w: 10, h: 0.18, fill: { color: AMBER } });
      es.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 5.445, w: 10, h: 0.18, fill: { color: AMBER } });
      es.addShape(pptx.shapes.OVAL, { x: 3.2, y: 0.7, w: 3.6, h: 3.6, fill: { color: "253A5E", transparency: 75 } });
      es.addText("END OF QUIZ", {
        x: 0.5, y: 1.6, w: 9, h: 1.6,
        fontSize: 52, color: WHITE, bold: true, align: "center", fontFace: "Cambria"
      });
      es.addText("Great job! Review your answers above.", {
        x: 0.5, y: 3.4, w: 9, h: 0.6,
        fontSize: 22, color: Q_HEADER, align: "center", fontFace: "Calibri"
      });

      const fileName = `${testTitle.replace(/[^a-z0-9]/gi, "_")}_Presentation.pptx`;
      await pptx.writeFile({ fileName });
      setStatus("✅ Presentation downloaded!");
      setStage("done");
    } catch (err) {
      setError("PPTX error: " + err.message);
      setStage("preview");
    }
  };

  // ── UI ─────────────────────────────────────────────────────────
  return (
    <div style={{ fontFamily: "'Segoe UI', system-ui, sans-serif", minHeight: "100vh", background: "linear-gradient(160deg, #111e30 0%, #162336 60%, #0f1e2c 100%)", color: "#ddeeff", padding: "0", fontSize: "16px" }}>
      {/* Header */}
      <div style={{ background: "rgba(15,28,44,0.97)", borderBottom: "2px solid #4EA8DE", padding: "18px 36px", display: "flex", alignItems: "center", gap: "14px", backdropFilter: "blur(10px)" }}>
        <div style={{ width: 46, height: 46, borderRadius: "50%", background: "#4EA8DE", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 24 }}>📋</div>
        <div>
          <div style={{ fontSize: 22, fontWeight: 700, color: "#fff", letterSpacing: 0.5 }}>QuizSlide Maker</div>
          <div style={{ fontSize: 12, color: "#7DBBD8", letterSpacing: 2 }}>UPLOAD · EXTRACT · PRESENT</div>
        </div>
        {questions.length > 0 && (
          <div style={{ marginLeft: "auto", background: "#1B3A6B", borderRadius: 20, padding: "8px 18px", fontSize: 14, color: "#4EA8DE", fontWeight: 600 }}>
            {questions.length} questions extracted
          </div>
        )}
      </div>

      <div style={{ maxWidth: 820, margin: "0 auto", padding: "44px 28px" }}>
        {/* Step Indicators */}
        <div style={{ display: "flex", justifyContent: "center", gap: 0, marginBottom: 40 }}>
          {["Upload", "Extract", "Preview", "Generate"].map((label, i) => {
            const stageMap = ["upload", "processing", "preview", "generating"];
            const doneIdx = ["upload", "processing", "preview", "generating", "done"].indexOf(stage);
            const isActive = stageMap[i] === stage || (stage === "done" && i === 3);
            const isDone = doneIdx > i;
            return (
              <div key={i} style={{ display: "flex", alignItems: "center" }}>
                <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 4 }}>
                  <div style={{ width: 34, height: 34, borderRadius: "50%", background: isDone ? "#2ECC71" : isActive ? "#F5A623" : "#1B3A6B", border: `2px solid ${isDone ? "#2ECC71" : isActive ? "#F5A623" : "#1B3A6B"}`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14, fontWeight: 700, color: isDone || isActive ? "#0D1B2A" : "#7A9BB5", transition: "all 0.3s" }}>
                    {isDone ? "✓" : i + 1}
                  </div>
                  <span style={{ fontSize: 11, color: isActive ? "#F5A623" : isDone ? "#2ECC71" : "#7A9BB5", fontWeight: isActive ? 700 : 400 }}>{label}</span>
                </div>
                {i < 3 && <div style={{ width: 60, height: 2, background: isDone ? "#2ECC71" : "#1B3A6B", marginBottom: 20, transition: "all 0.3s" }} />}
              </div>
            );
          })}
        </div>

        {/* Error Banner */}
        {error && (
          <div style={{ background: "rgba(220,38,38,0.12)", border: "1px solid #dc2626", borderRadius: 10, padding: "14px 18px", marginBottom: 16, color: "#fca5a5", fontSize: 13, wordBreak: "break-word" }}>
            <div style={{ fontWeight: 700, marginBottom: 4 }}>⚠️ Error</div>
            <div>{error}</div>
          </div>
        )}

        {/* Debug Log */}
        {debugLog.length > 0 && (
          <div style={{ background: "rgba(0,0,0,0.4)", border: "1px solid #1B3A6B", borderRadius: 10, padding: "12px 16px", marginBottom: 16, fontSize: 11, fontFamily: "monospace", color: "#7A9BB5", maxHeight: 140, overflowY: "auto" }}>
            <div style={{ color: "#F5A623", fontWeight: 700, marginBottom: 6 }}>DEBUG LOG</div>
            {debugLog.map((l, i) => <div key={i} style={{ marginBottom: 2 }}>› {l}</div>)}
          </div>
        )}

        {/* ── UPLOAD STAGE ── */}
        {stage === "upload" && (
          <div>
            <div style={{ textAlign: "center", marginBottom: 32 }}>
              <h1 style={{ fontSize: 32, fontWeight: 700, color: "#fff", margin: "0 0 10px" }}>Upload Your Test Paper</h1>
              <p style={{ color: "#7DBBD8", fontSize: 16, margin: 0 }}>Supports PDF, DOCX, or TXT test formats. No API key required.</p>
            </div>

            <div
              onDrop={handleDrop}
              onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onClick={() => fileInputRef.current?.click()}
              style={{ border: `2px dashed ${dragOver ? "#F5A623" : file ? "#2ECC71" : "#1B3A6B"}`, borderRadius: 16, padding: "52px 32px", textAlign: "center", cursor: "pointer", background: dragOver ? "rgba(245,166,35,0.05)" : file ? "rgba(46,204,113,0.05)" : "rgba(27,58,107,0.15)", transition: "all 0.25s", marginBottom: 28 }}
            >
              <input ref={fileInputRef} type="file" accept=".pdf,.txt,.docx,.doc" style={{ display: "none" }} onChange={(e) => readFile(e.target.files[0])} />
              <div style={{ fontSize: 58, marginBottom: 14 }}>{file ? "📄" : "📁"}</div>
              {file ? (
                <>
                  <div style={{ fontSize: 20, fontWeight: 600, color: "#3DD68C", marginBottom: 8 }}>{file.name}</div>
                  <div style={{ fontSize: 14, color: "#7DBBD8" }}>{(file.size / 1024).toFixed(1)} KB · Click to change file</div>
                {fileData && <div style={{ marginTop: 6, fontSize: 14, color: "#3DD68C", fontWeight: 600 }}>✅ File parsed and ready</div>}
                {!fileData && <div style={{ marginTop: 6, fontSize: 14, color: "#F0A500" }}>⏳ Parsing file...</div>}
                </>
              ) : (
                <>
                  <div style={{ fontSize: 19, fontWeight: 600, color: "#E8F4FD", marginBottom: 8 }}>Drop your test paper here</div>
                  <div style={{ fontSize: 15, color: "#7DBBD8" }}>or click to browse · PDF, DOCX, TXT supported</div>
                </>
              )}
            </div>

            {/* Randomize toggle */}
            <div style={{ background: "rgba(27,58,107,0.3)", borderRadius: 12, padding: "18px 22px", marginBottom: 28, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
              <div>
                <div style={{ fontWeight: 600, color: "#E8F4FD", fontSize: 16 }}>🔀 Randomize Questions</div>
                <div style={{ color: "#7DBBD8", fontSize: 14, marginTop: 4 }}>Shuffle question order in the presentation</div>
              </div>
              <div
                onClick={() => setRandomize(!randomize)}
                style={{ width: 52, height: 28, borderRadius: 14, background: randomize ? "#F5A623" : "#1B3A6B", cursor: "pointer", position: "relative", transition: "all 0.3s" }}
              >
                <div style={{ position: "absolute", top: 3, left: randomize ? 26 : 3, width: 22, height: 22, borderRadius: "50%", background: "#fff", transition: "all 0.3s", boxShadow: "0 1px 4px rgba(0,0,0,0.3)" }} />
              </div>
            </div>

            <button
              onClick={extractQuestions}
              disabled={!file}
              style={{ width: "100%", padding: "18px", borderRadius: 14, border: "none", background: file ? "linear-gradient(135deg, #4EA8DE, #2a88c8)" : "#1B3A6B", color: file ? "#fff" : "#7A9BB5", fontSize: 18, fontWeight: 700, cursor: file ? "pointer" : "not-allowed", transition: "all 0.2s", letterSpacing: 0.5 }}
            >
              {status === "Reading file..." ? "⏳ Reading file..." : "🔍 Extract Questions"}
            </button>
          </div>
        )}

        {/* ── PROCESSING STAGE ── */}
        {stage === "processing" && (
          <div style={{ textAlign: "center", padding: "60px 0" }}>
            <div style={{ fontSize: 64, marginBottom: 20, animation: "spin 1.5s linear infinite" }}>⚙️</div>
            <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } } @keyframes pulse { 0%,100% { opacity:1; } 50% { opacity:0.4; } }`}</style>
            <h2 style={{ fontSize: 26, color: "#fff", marginBottom: 12 }}>Processing Test Paper</h2>
            <p style={{ color: "#7A9BB5", fontSize: 15, animation: "pulse 2s ease-in-out infinite" }}>{status}</p>
            <div style={{ marginTop: 30, display: "flex", justifyContent: "center", gap: 8 }}>
              {[0,1,2].map(i => <div key={i} style={{ width: 10, height: 10, borderRadius: "50%", background: "#F5A623", animation: `pulse 1.2s ease-in-out ${i * 0.2}s infinite` }} />)}
            </div>
          </div>
        )}

        {/* ── PREVIEW STAGE ── */}
        {stage === "preview" && (
          <div>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 24 }}>
              <div>
                <h2 style={{ fontSize: 24, fontWeight: 700, color: "#fff", margin: "0 0 6px" }}>📋 {testTitle}</h2>
                <p style={{ color: "#7DBBD8", fontSize: 15, margin: 0 }}>{questions.length} questions extracted — review before generating</p>
              </div>
              <button onClick={() => setStage("upload")} style={{ background: "transparent", border: "1px solid #1B3A6B", borderRadius: 8, padding: "8px 14px", color: "#7A9BB5", cursor: "pointer", fontSize: 13 }}>← Back</button>
            </div>

            {/* Randomize toggle */}
            <div style={{ background: "rgba(27,58,107,0.3)", borderRadius: 10, padding: "14px 18px", marginBottom: 20, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
              <div style={{ fontWeight: 600, color: "#E8F4FD", fontSize: 14 }}>🔀 Randomize in presentation</div>
              <div onClick={() => setRandomize(!randomize)} style={{ width: 52, height: 28, borderRadius: 14, background: randomize ? "#F5A623" : "#1B3A6B", cursor: "pointer", position: "relative", transition: "all 0.3s" }}>
                <div style={{ position: "absolute", top: 3, left: randomize ? 26 : 3, width: 22, height: 22, borderRadius: "50%", background: "#fff", transition: "all 0.3s" }} />
              </div>
            </div>

            {/* Questions list */}
            <div style={{ maxHeight: 380, overflowY: "auto", marginBottom: 24, display: "flex", flexDirection: "column", gap: 10, paddingRight: 4 }}>
              {questions.map((q, i) => (
                <div key={i} style={{ background: "rgba(27,58,107,0.25)", border: "1px solid #1B3A6B", borderRadius: 10, overflow: "hidden" }}>
                  <div
                    onClick={() => setExpandedQ(expandedQ === i ? null : i)}
                    style={{ display: "flex", alignItems: "flex-start", gap: 14, padding: "14px 18px", cursor: "pointer" }}
                  >
                    <div style={{ width: 30, height: 30, borderRadius: "50%", background: "#F5A623", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13, fontWeight: 700, color: "#0D1B2A", flexShrink: 0, marginTop: 1 }}>{i + 1}</div>
                    <div style={{ flex: 1 }}>
                      <div style={{ color: "#E8F4FD", fontSize: 15, lineHeight: 1.6 }}>{q.question}</div>
                      {q.choices?.length > 0 && <div style={{ marginTop: 5, fontSize: 13, color: "#7DBBD8" }}>{q.choices.length} choices</div>}
                    </div>
                    <div style={{ color: "#7A9BB5", fontSize: 18, transform: expandedQ === i ? "rotate(180deg)" : "none", transition: "transform 0.2s" }}>▾</div>
                  </div>
                  {expandedQ === i && (
                    <div style={{ padding: "0 18px 16px 62px", borderTop: "1px solid rgba(27,58,107,0.5)" }}>
                      {q.choices?.length > 0 && (
                        <div style={{ marginTop: 10, display: "flex", flexDirection: "column", gap: 5 }}>
                          {q.choices.map((ch, ci) => (
                            <div key={ci} style={{ fontSize: 14, color: "#B8D4E8", lineHeight: 1.6, padding: "4px 8px", background: "rgba(27,58,107,0.3)", borderRadius: 5 }}>{ch}</div>
                          ))}
                        </div>
                      )}
                      <div style={{ marginTop: 10, padding: "10px 12px", background: "rgba(46,204,113,0.1)", borderRadius: 7, borderLeft: "3px solid #2ECC71" }}>
                        <div style={{ fontSize: 11, color: "#7AC99A", fontWeight: 700, marginBottom: 3 }}>ANSWER</div>
                        <div style={{ fontSize: 15, color: "#E8F4FD" }}>{q.answer}</div>
                      </div>
                      {q.rationale && (
                        <div style={{ marginTop: 8, padding: "10px 12px", background: "rgba(245,166,35,0.08)", borderRadius: 7, borderLeft: "3px solid #F5A623" }}>
                          <div style={{ fontSize: 11, color: "#F5A623", fontWeight: 700, marginBottom: 3 }}>RATIONALE</div>
                          <div style={{ fontSize: 14, color: "#B8D4E8", lineHeight: 1.6 }}>{q.rationale}</div>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              ))}
            </div>

            <button
              onClick={generatePPTX}
              style={{ width: "100%", padding: "18px", borderRadius: 14, border: "none", background: "linear-gradient(135deg, #3DD68C, #1db85e)", color: "#0A1F0F", fontSize: 18, fontWeight: 700, cursor: "pointer", letterSpacing: 0.5 }}
            >
              🎯 Generate PowerPoint Presentation
            </button>
          </div>
        )}

        {/* ── GENERATING STAGE ── */}
        {stage === "generating" && (
          <div style={{ textAlign: "center", padding: "60px 0" }}>
            <div style={{ fontSize: 64, marginBottom: 20 }}>🎬</div>
            <h2 style={{ fontSize: 24, color: "#fff", marginBottom: 10 }}>Building Slides</h2>
            <p style={{ color: "#7A9BB5", fontSize: 15 }}>{status}</p>
            <div style={{ marginTop: 24, background: "rgba(27,58,107,0.4)", borderRadius: 8, height: 6, overflow: "hidden" }}>
              <div style={{ height: "100%", width: "70%", background: "linear-gradient(90deg, #F5A623, #2ECC71)", borderRadius: 8, animation: "pulse 1.5s ease-in-out infinite" }} />
            </div>
          </div>
        )}

        {/* ── DONE STAGE ── */}
        {stage === "done" && (
          <div style={{ textAlign: "center", padding: "50px 0" }}>
            <div style={{ fontSize: 72, marginBottom: 16 }}>🎉</div>
            <h2 style={{ fontSize: 32, fontWeight: 700, color: "#fff", marginBottom: 12 }}>Presentation Ready!</h2>
            <p style={{ color: "#3DD68C", fontSize: 17, marginBottom: 34 }}>Your PPTX file has been downloaded automatically.</p>
            <div style={{ background: "rgba(27,58,107,0.3)", borderRadius: 12, padding: "22px", marginBottom: 28, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
              {[["📄", "Slides", `${questions.length * 2 + 2} total`], ["❓", "Questions", questions.length], ["✅", "Answers", questions.length], ["🔀", "Order", randomize ? "Randomized" : "Original"]].map(([icon, label, val]) => (
                <div key={label} style={{ background: "rgba(13,27,42,0.5)", borderRadius: 8, padding: "14px", textAlign: "center" }}>
                  <div style={{ fontSize: 26 }}>{icon}</div>
                  <div style={{ fontSize: 12, color: "#7A9BB5", marginTop: 4 }}>{label}</div>
                  <div style={{ fontSize: 18, fontWeight: 700, color: "#F5A623", marginTop: 2 }}>{val}</div>
                </div>
              ))}
            </div>
            <div style={{ display: "flex", gap: 12, justifyContent: "center" }}>
              <button onClick={() => { setStage("preview"); }} style={{ padding: "12px 24px", borderRadius: 10, border: "1px solid #2ECC71", background: "transparent", color: "#2ECC71", fontSize: 15, fontWeight: 600, cursor: "pointer" }}>
                🔄 Re-generate
              </button>
              <button onClick={() => { setStage("upload"); setFile(null); setFileData(null); setQuestions([]); setError(""); }} style={{ padding: "12px 24px", borderRadius: 10, border: "none", background: "linear-gradient(135deg, #F5A623, #e8940f)", color: "#0D1B2A", fontSize: 15, fontWeight: 700, cursor: "pointer" }}>
                📁 New Test Paper
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
