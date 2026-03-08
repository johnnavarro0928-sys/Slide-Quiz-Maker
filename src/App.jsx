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
const DEPED = {
  blue: "#0038A8",
  red: "#CE1126",
  yellow: "#FCD116",
  text: "#122347",
  muted: "#5C6B8A",
  bg: "#E6ECF4",
  surface: "#ECF1F8",
  surfaceAlt: "#E2E9F4",
  border: "#CFD9E8",
  success: "#1E8E5A",
  error: "#B21F2D"
};

const neuRaised = (radius = 16) => ({
  borderRadius: radius,
  background: DEPED.surface,
  boxShadow: "10px 10px 22px #c5ceda, -10px -10px 22px #ffffff"
});

const neuInset = (radius = 16) => ({
  borderRadius: radius,
  background: DEPED.surface,
  boxShadow: "inset 8px 8px 14px #c9d2de, inset -8px -8px 14px #ffffff"
});

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

    const simpleLetter = answer.match(/^\(?([A-H])\)?$/i);
    const prefixedLetter = answer.match(/^\(?([A-H])\)?\s*[\).:\-]/i);
    const optionChoiceLetter = answer.match(/\b(?:option|choice)\s*([A-H])\b/i);
    const letter = (simpleLetter?.[1] || prefixedLetter?.[1] || optionChoiceLetter?.[1] || "").toUpperCase();

    if (letter && Array.isArray(choices) && choices.length > 0) {
      const found = choices.find((choice) => choice.toUpperCase().startsWith(`${letter}.`));
      return found || `${letter}.`;
    }

    if (letter) return `${letter}.`;
    return answer;
  }, []);

  const parseAnswerKeyMap = useCallback((keyLines) => {
    const keyMap = new Map();

    keyLines.forEach((rawLine) => {
      const line = rawLine.trim();
      if (!line) return;

      const strictRe = /(\d{1,4})\s*[\).:\-]\s*([A-H])\b/gi;
      let strictMatch = strictRe.exec(line);
      let foundStrict = false;
      while (strictMatch) {
        keyMap.set(Number(strictMatch[1]), strictMatch[2].toUpperCase());
        foundStrict = true;
        strictMatch = strictRe.exec(line);
      }
      if (foundStrict) return;

      const looseRe = /(\d{1,4})\s+([A-H])\b/gi;
      let looseMatch = looseRe.exec(line);
      while (looseMatch) {
        keyMap.set(Number(looseMatch[1]), looseMatch[2].toUpperCase());
        looseMatch = looseRe.exec(line);
      }
    });

    return keyMap;
  }, []);

  const parseRationaleKeyMap = useCallback((keyLines) => {
    const rationaleMap = new Map();
    let currentNumber = null;

    keyLines.forEach((rawLine) => {
      const line = rawLine.trim();
      if (!line) return;

      const answerPlusRationaleMatch = line.match(/^(\d{1,4})\s*[\).:\-]\s*[A-H]\b\s*(.*)$/i);
      if (answerPlusRationaleMatch) {
        currentNumber = Number(answerPlusRationaleMatch[1]);
        let rationaleText = (answerPlusRationaleMatch[2] || "").trim();
        rationaleText = rationaleText.replace(/^[\-:–—]\s*/, "");
        rationaleText = rationaleText.replace(/^\((.*)\)$/s, "$1").trim();
        if (rationaleText) {
          rationaleMap.set(currentNumber, rationaleText);
        }
        return;
      }

      const numberedMatch = line.match(/^(\d{1,4})\s*[\).:\-]\s*(.+)$/);
      if (numberedMatch) {
        currentNumber = Number(numberedMatch[1]);
        rationaleMap.set(currentNumber, numberedMatch[2].trim());
        return;
      }

      if (currentNumber !== null) {
        const prev = rationaleMap.get(currentNumber) || "";
        rationaleMap.set(currentNumber, `${prev} ${line}`.replace(/\s+/g, " ").trim());
      }
    });

    return rationaleMap;
  }, []);

  const parseQuestionsFromText = useCallback((rawText) => {
    const text = (rawText || "").replace(/\r/g, "");
    const allLines = text.split("\n").map((line) => line.trimEnd());
    const questionRe = /^\s*(?:q(?:uestion)?\s*)?(\d{1,4})\s*[\).:\-]\s*(.+)$/i;
    const choiceRe = /^\s*([A-H])\s*[\).:\-]\s*(.+)$/i;
    const answerLabelRe = /^\s*(?:answer|correct answer|answer key|key answer|ans(?:wer)?(?:\s*key)?)\s*[:\-]\s*(.*)$/i;
    const answerKeyHeadingRe = /^\s*(?:answer\s*key(?:\s*&\s*rationale)?|key\s*answers?|answer\s*keys)\s*[:\-]?\s*$/i;
    const answerKeyHeadingWithContentRe = /^\s*(?:answer\s*key(?:\s*&\s*rationale)?|key\s*answers?|answer\s*keys)\s*[:\-]\s*(.+)$/i;
    const rationaleKeyHeadingRe = /^\s*(?:rationale(?:\s*key)?|rationales|explanation(?:s)?(?:\s*key)?|reason(?:ing)?(?:\s*key)?)\s*[:\-]?\s*$/i;
    const rationaleKeyHeadingWithContentRe = /^\s*(?:rationale(?:\s*key)?|rationales|explanation(?:s)?(?:\s*key)?|reason(?:ing)?(?:\s*key)?)\s*[:\-]\s*(.+)$/i;
    const rationaleRe = /^\s*(?:rationale|explanation|reason(?:ing)?)\s*[:\-]\s*(.*)$/i;
    const answerSectionLines = [];
    const rationaleSectionLines = [];

    const firstQuestionIdx = allLines.findIndex((line) => questionRe.test(line));
    let title = "Test";
    if (firstQuestionIdx > 0) {
      const maybeTitle = allLines.slice(0, firstQuestionIdx).find((line) => line.trim().length > 0);
      if (maybeTitle) {
        title = maybeTitle.replace(/^(?:title|quiz|test)\s*[:\-]\s*/i, "").trim() || "Test";
      }
    }

    const extracted = [];
    let current = null;
    let section = "question";
    let mode = "questions";

    const finalizeCurrent = () => {
      if (!current || !current.question.trim()) return;
      extracted.push({
        number: current.number || extracted.length + 1,
        question: current.question.replace(/\s+/g, " ").trim(),
        choices: current.choices.map((choice) => choice.replace(/\s+/g, " ").trim()),
        answer: normalizeAnswer(current.answer, current.choices),
        rationale: current.rationale.replace(/\s+/g, " ").trim()
      });
      current = null;
      section = "question";
    };

    allLines.forEach((rawLine) => {
      const line = rawLine.trim();
      if (!line) return;

      if (answerKeyHeadingRe.test(line)) {
        finalizeCurrent();
        mode = /rationale/i.test(line) ? "answerAndRationaleKey" : "answerKey";
        return;
      }

      const answerHeadingWithContent = line.match(answerKeyHeadingWithContentRe);
      if (answerHeadingWithContent) {
        const payload = answerHeadingWithContent[1].trim();
        if (!current || /\d{1,4}\s*[\).:\-]\s*[A-H]/i.test(payload)) {
          finalizeCurrent();
          mode = /rationale/i.test(line) ? "answerAndRationaleKey" : "answerKey";
          answerSectionLines.push(payload);
          if (/rationale/i.test(line)) rationaleSectionLines.push(payload);
          return;
        }
      }

      if (rationaleKeyHeadingRe.test(line)) {
        finalizeCurrent();
        mode = "rationaleKey";
        return;
      }

      const rationaleHeadingWithContent = line.match(rationaleKeyHeadingWithContentRe);
      if (rationaleHeadingWithContent) {
        const payload = rationaleHeadingWithContent[1].trim();
        if (!current || /\d{1,4}\s*[\).:\-]/.test(payload)) {
          finalizeCurrent();
          mode = "rationaleKey";
          rationaleSectionLines.push(payload);
          return;
        }
      }

      if (mode === "answerKey") {
        answerSectionLines.push(line);
        return;
      }

      if (mode === "answerAndRationaleKey") {
        answerSectionLines.push(line);
        rationaleSectionLines.push(line);
        return;
      }

      if (mode === "rationaleKey") {
        rationaleSectionLines.push(line);
        return;
      }

      const qMatch = line.match(questionRe);
      if (qMatch) {
        mode = "questions";
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

      const answerMatch = line.match(answerLabelRe);
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
    const answerKeyMap = parseAnswerKeyMap(answerSectionLines);
    const rationaleKeyMap = parseRationaleKeyMap(rationaleSectionLines);

    extracted.forEach((q) => {
      const fallbackLetter = answerKeyMap.get(q.number);
      if (fallbackLetter && (!q.answer || q.answer === "Not provided")) {
        q.answer = normalizeAnswer(fallbackLetter, q.choices);
      }

      const fallbackRationale = rationaleKeyMap.get(q.number);
      if ((!q.rationale || q.rationale.length === 0) && fallbackRationale) {
        q.rationale = fallbackRationale;
      }

      if (!q.rationale || q.rationale.length === 0) {
        q.rationale = "No rationale provided.";
      }
    });

    if (extracted.length === 0) {
      throw new Error(
        "No questions were detected. Please use a format like '1. ...', choices 'A. ...', and include 'Answer:' / 'Rationale:'."
      );
    }

    return { title, questions: extracted };
  }, [normalizeAnswer, parseAnswerKeyMap, parseRationaleKeyMap]);

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
      pptx.author = "QuizSlide Maker";
      pptx.company = "DepEd";
      pptx.subject = "Assessment Review";
      pptx.title = testTitle;
      const PPT = {
        depedBlue: "0038A8",
        depedRed: "CE1126",
        depedYellow: "FCD116",
        slate: "23324A",
        ink: "1B2430",
        muted: "5D6A7B",
        paper: "F8FAFD",
        panel: "FFFFFF",
        border: "D8E0EA",
        accentFill: "EAF0FF",
        goodFill: "EAF7EE",
        goodBorder: "66A97A"
      };

      const sanitizeText = (value, fallback = "") => {
        if (typeof value !== "string") return fallback;
        const cleaned = value.replace(/\s+/g, " ").trim();
        return cleaned || fallback;
      };

      const wrapLines = (text, maxCharsPerLine) => {
        const words = sanitizeText(text).split(" ").filter(Boolean);
        const lines = [];
        let current = "";
        words.forEach((word) => {
          if (!current) {
            current = word;
            return;
          }
          const candidate = `${current} ${word}`;
          if (candidate.length <= maxCharsPerLine) {
            current = candidate;
          } else {
            lines.push(current);
            current = word;
          }
        });
        if (current) lines.push(current);
        return lines.length ? lines : [""];
      };

      const estimateLines = (text, boxW, fontSize) => {
        const charsPerLine = Math.max(16, Math.floor((boxW * 72) / (fontSize * 0.54)));
        const chunks = String(text || "").split(/\n+/).flatMap((seg) => wrapLines(seg, charsPerLine));
        return Math.max(1, chunks.length);
      };

      const linesToInches = (lines, fontSize, lineSpacing = 1.2) => (lines * fontSize * lineSpacing) / 72;

      const fitFontSize = (text, boxW, maxH, sizes, lineSpacing = 1.2) => {
        for (const size of sizes) {
          const lines = estimateLines(text, boxW, size);
          if (linesToInches(lines, size, lineSpacing) <= maxH) return size;
        }
        return sizes[sizes.length - 1];
      };

      const clampTextToBox = (text, boxW, boxH, fontSize, lineSpacing = 1.2) => {
        const source = sanitizeText(text, "");
        const charsPerLine = Math.max(16, Math.floor((boxW * 72) / (fontSize * 0.54)));
        const maxLines = Math.max(1, Math.floor((boxH * 72) / (fontSize * lineSpacing)));
        const lines = source.split(/\n+/).flatMap((segment) => wrapLines(segment, charsPerLine));
        if (lines.length <= maxLines) return lines.join("\n");
        const trimmed = lines.slice(0, maxLines);
        trimmed[maxLines - 1] = `${trimmed[maxLines - 1].slice(0, Math.max(4, charsPerLine - 2)).trimEnd()}…`;
        return trimmed.join("\n");
      };

      const stripChoicePrefix = (choiceText) => sanitizeText(choiceText).replace(/^[A-Ha-h][.)]\s*/, "");

      const totalSlides = finalQs.length * 2 + 2;
      let slideNo = 1;
      const titleForFooter = sanitizeText(testTitle, "Quiz Presentation").slice(0, 56);

      const addFooter = (slide) => {
        slide.addShape(pptx.shapes.RECTANGLE, {
          x: 0,
          y: 5.24,
          w: 10,
          h: 0.015,
          fill: { color: PPT.border },
          line: { color: PPT.border, width: 0 }
        });
        slide.addText(`${titleForFooter}   |   ${slideNo} / ${totalSlides}`, {
          x: 0.45,
          y: 5.29,
          w: 9.1,
          h: 0.2,
          fontFace: "Calibri",
          fontSize: 10,
          color: PPT.muted,
          align: "right"
        });
        slideNo += 1;
      };

      // ── TITLE SLIDE ─────────────────────────────────────────────
      const titleSlide = pptx.addSlide();
      titleSlide.background = { color: PPT.paper };
      titleSlide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.22, fill: { color: PPT.depedBlue }, line: { color: PPT.depedBlue, width: 0 } });
      titleSlide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0.22, w: 10, h: 0.04, fill: { color: PPT.depedYellow }, line: { color: PPT.depedYellow, width: 0 } });
      titleSlide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 5.365, w: 10, h: 0.26, fill: { color: PPT.depedRed }, line: { color: PPT.depedRed, width: 0 } });
      titleSlide.addShape(pptx.shapes.OVAL, { x: 6.9, y: -0.55, w: 4.0, h: 4.0, fill: { color: "E7EEFC" }, line: { color: "E7EEFC", width: 0 } });
      titleSlide.addShape(pptx.shapes.OVAL, { x: -1.2, y: 3.1, w: 2.8, h: 2.8, fill: { color: "E7EEFC" }, line: { color: "E7EEFC", width: 0 } });
      titleSlide.addText("QUIZ PRESENTATION", {
        x: 0.7,
        y: 1.25,
        w: 8.6,
        h: 0.36,
        fontFace: "Calibri",
        fontSize: 15,
        color: PPT.depedBlue,
        bold: true,
        charSpacing: 4,
        align: "center"
      });
      const titleText = clampTextToBox(testTitle, 8.6, 1.6, 44, 1.08);
      titleSlide.addText(titleText, {
        x: 0.7,
        y: 1.72,
        w: 8.6,
        h: 1.8,
        fontFace: "Cambria",
        fontSize: 44,
        color: PPT.slate,
        bold: true,
        align: "center",
        valign: "mid",
        lineSpacingMultiple: 1.08
      });
      titleSlide.addShape(pptx.shapes.RECTANGLE, {
        x: 2.1,
        y: 4.08,
        w: 5.8,
        h: 0.74,
        fill: { color: PPT.accentFill },
        line: { color: PPT.border, width: 1.2 }
      });
      titleSlide.addText(`${finalQs.length} Questions  •  ${randomize ? "Randomized order" : "Original order"}`, {
        x: 2.1,
        y: 4.26,
        w: 5.8,
        h: 0.34,
        fontFace: "Calibri",
        fontSize: 17,
        color: PPT.slate,
        align: "center",
        bold: true
      });
      addFooter(titleSlide);

      // ── QUESTION + ANSWER SLIDES ──────────────────────────────
      finalQs.forEach((rawQ, idx) => {
        const qNum = idx + 1;
        const question = sanitizeText(rawQ.question, "Question text not provided.");
        const choices = Array.isArray(rawQ.choices) ? rawQ.choices.filter(Boolean).map((c) => sanitizeText(c)) : [];
        const answer = sanitizeText(rawQ.answer, "Not provided");
        const rationale = sanitizeText(rawQ.rationale, "No rationale provided.");
        const hasChoices = choices.length > 0;

        // ── QUESTION SLIDE ──────────────────────────────────────
        const qSlide = pptx.addSlide();
        qSlide.background = { color: PPT.paper };
        qSlide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.88, fill: { color: PPT.depedBlue }, line: { color: PPT.depedBlue, width: 0 } });
        qSlide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0.88, w: 10, h: 0.035, fill: { color: PPT.depedYellow }, line: { color: PPT.depedYellow, width: 0 } });
        qSlide.addShape(pptx.shapes.OVAL, { x: 0.38, y: 0.14, w: 0.56, h: 0.56, fill: { color: "FFFFFF" }, line: { color: "FFFFFF", width: 0 } });
        qSlide.addText(String(qNum), {
          x: 0.38,
          y: 0.14,
          w: 0.56,
          h: 0.56,
          fontFace: "Calibri",
          fontSize: 20,
          color: PPT.depedBlue,
          bold: true,
          align: "center",
          valign: "mid",
          margin: 0
        });
        qSlide.addText(`QUESTION ${qNum}`, {
          x: 1.06,
          y: 0.23,
          w: 4.4,
          h: 0.36,
          fontFace: "Calibri",
          fontSize: 19,
          color: "FFFFFF",
          bold: true,
          charSpacing: 1
        });
        qSlide.addShape(pptx.shapes.RECTANGLE, {
          x: 0.45,
          y: 1.0,
          w: 9.1,
          h: 4.12,
          fill: { color: PPT.panel },
          line: { color: PPT.border, width: 1.25 }
        });

        const questionBoxX = 0.75;
        const questionBoxY = 1.28;
        const questionBoxW = 8.5;
        const rows = hasChoices ? Math.ceil(choices.length / (choices.length > 2 ? 2 : 1)) : 0;
        const estimatedChoiceBlock = hasChoices ? Math.max(1.3, rows * 0.72 + (rows - 1) * 0.12) : 0;
        const questionBoxH = hasChoices ? Math.max(1.0, 3.56 - estimatedChoiceBlock) : 3.3;
        const questionFont = fitFontSize(question, questionBoxW, questionBoxH, [34, 31, 28, 25, 22, 20, 18], 1.24);
        const clampedQuestion = clampTextToBox(question, questionBoxW, questionBoxH, questionFont, 1.24);

        qSlide.addText(clampedQuestion, {
          x: questionBoxX,
          y: questionBoxY,
          w: questionBoxW,
          h: questionBoxH,
          fontFace: "Calibri",
          fontSize: questionFont,
          color: PPT.ink,
          bold: true,
          valign: "top",
          wrap: true,
          lineSpacingMultiple: 1.24
        });

        if (hasChoices) {
          const cols = choices.length > 2 ? 2 : 1;
          const gap = 0.12;
          const choiceStartY = questionBoxY + questionBoxH + 0.18;
          const choiceAreaH = 5.0 - choiceStartY;
          const cardW = cols === 2 ? (questionBoxW - gap) / 2 : questionBoxW;
          const cardH = Math.max(0.62, (choiceAreaH - (rows - 1) * gap) / rows);
          const longestChoice = choices.reduce((a, b) => (a.length > b.length ? a : b), "");
          const choiceFont = fitFontSize(stripChoicePrefix(longestChoice), cardW - 0.6, cardH - 0.12, [19, 18, 17, 16, 15, 14, 13], 1.15);

          choices.forEach((choice, i) => {
            const col = Math.floor(i / rows);
            const row = i % rows;
            const x = questionBoxX + col * (cardW + gap);
            const y = choiceStartY + row * (cardH + gap);
            const letter = String.fromCharCode(65 + i);
            const choiceText = stripChoicePrefix(choice);
            const renderedChoice = clampTextToBox(choiceText, cardW - 0.6, cardH - 0.1, choiceFont, 1.15);

            qSlide.addShape(pptx.shapes.RECTANGLE, {
              x,
              y,
              w: cardW,
              h: cardH,
              fill: { color: "F7FAFF" },
              line: { color: PPT.border, width: 1 }
            });
            qSlide.addShape(pptx.shapes.OVAL, {
              x: x + 0.1,
              y: y + (cardH - 0.26) / 2,
              w: 0.26,
              h: 0.26,
              fill: { color: PPT.depedBlue },
              line: { color: PPT.depedBlue, width: 0 }
            });
            qSlide.addText(letter, {
              x: x + 0.1,
              y: y + (cardH - 0.26) / 2,
              w: 0.26,
              h: 0.26,
              fontFace: "Calibri",
              fontSize: choiceFont - 2,
              color: "FFFFFF",
              bold: true,
              align: "center",
              valign: "mid",
              margin: 0
            });
            qSlide.addText(renderedChoice, {
              x: x + 0.42,
              y: y + 0.04,
              w: cardW - 0.48,
              h: cardH - 0.08,
              fontFace: "Calibri",
              fontSize: choiceFont,
              color: PPT.ink,
              valign: "mid",
              wrap: true,
              lineSpacingMultiple: 1.15
            });
          });
        }
        addFooter(qSlide);

        // ── ANSWER SLIDE ────────────────────────────────────────
        const answerSlide = pptx.addSlide();
        answerSlide.background = { color: PPT.paper };
        answerSlide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.88, fill: { color: PPT.depedRed }, line: { color: PPT.depedRed, width: 0 } });
        answerSlide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0.88, w: 10, h: 0.035, fill: { color: PPT.depedYellow }, line: { color: PPT.depedYellow, width: 0 } });
        answerSlide.addText(`ANSWER KEY  •  QUESTION ${qNum}`, {
          x: 0.6,
          y: 0.26,
          w: 8.8,
          h: 0.35,
          fontFace: "Calibri",
          fontSize: 18,
          color: "FFFFFF",
          bold: true,
          charSpacing: 1
        });
        answerSlide.addShape(pptx.shapes.RECTANGLE, {
          x: 0.45,
          y: 1.0,
          w: 9.1,
          h: 4.12,
          fill: { color: PPT.panel },
          line: { color: PPT.border, width: 1.25 }
        });

        const promptPreview = clampTextToBox(question, 8.4, 0.62, 14, 1.18);
        answerSlide.addText(promptPreview, {
          x: 0.75,
          y: 1.18,
          w: 8.4,
          h: 0.62,
          fontFace: "Calibri",
          fontSize: 14,
          color: PPT.muted,
          italic: true,
          valign: "top",
          wrap: true
        });

        const answerBoxY = 1.9;
        const answerBoxH = 1.05;
        const answerFont = fitFontSize(answer, 8.35, 0.6, [30, 27, 24, 22, 20, 18], 1.2);
        const clampedAnswer = clampTextToBox(answer, 8.35, 0.6, answerFont, 1.2);
        answerSlide.addShape(pptx.shapes.RECTANGLE, {
          x: 0.75,
          y: answerBoxY,
          w: 8.5,
          h: answerBoxH,
          fill: { color: PPT.goodFill },
          line: { color: PPT.goodBorder, width: 1.4 }
        });
        answerSlide.addText("CORRECT ANSWER", {
          x: 0.95,
          y: answerBoxY + 0.08,
          w: 4.3,
          h: 0.24,
          fontFace: "Calibri",
          fontSize: 12,
          bold: true,
          color: "2A5B37",
          charSpacing: 1.4
        });
        answerSlide.addText(clampedAnswer, {
          x: 0.95,
          y: answerBoxY + 0.34,
          w: 8.1,
          h: 0.55,
          fontFace: "Calibri",
          fontSize: answerFont,
          color: "1B4028",
          bold: true,
          valign: "mid",
          wrap: true,
          lineSpacingMultiple: 1.2
        });

        const rationaleBoxY = 3.14;
        const rationaleBoxH = 1.8;
        const rationaleFont = fitFontSize(rationale, 8.2, rationaleBoxH - 0.32, [20, 18, 17, 16, 15, 14, 13], 1.3);
        const clampedRationale = clampTextToBox(rationale, 8.2, rationaleBoxH - 0.32, rationaleFont, 1.3);
        answerSlide.addText("RATIONALE", {
          x: 0.75,
          y: rationaleBoxY - 0.26,
          w: 2.8,
          h: 0.22,
          fontFace: "Calibri",
          fontSize: 12,
          bold: true,
          color: PPT.depedBlue,
          charSpacing: 1.3
        });
        answerSlide.addShape(pptx.shapes.RECTANGLE, {
          x: 0.75,
          y: rationaleBoxY,
          w: 8.5,
          h: rationaleBoxH,
          fill: { color: "F9FBFE" },
          line: { color: PPT.border, width: 1.1 }
        });
        answerSlide.addText(clampedRationale, {
          x: 0.95,
          y: rationaleBoxY + 0.14,
          w: 8.1,
          h: rationaleBoxH - 0.22,
          fontFace: "Calibri",
          fontSize: rationaleFont,
          color: PPT.ink,
          valign: "top",
          wrap: true,
          lineSpacingMultiple: 1.3
        });
        addFooter(answerSlide);
      });

      // ── END SLIDE ──────────────────────────────────────────────
      const endSlide = pptx.addSlide();
      endSlide.background = { color: PPT.paper };
      endSlide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.22, fill: { color: PPT.depedBlue }, line: { color: PPT.depedBlue, width: 0 } });
      endSlide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0.22, w: 10, h: 0.04, fill: { color: PPT.depedYellow }, line: { color: PPT.depedYellow, width: 0 } });
      endSlide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 5.365, w: 10, h: 0.26, fill: { color: PPT.depedRed }, line: { color: PPT.depedRed, width: 0 } });
      endSlide.addShape(pptx.shapes.OVAL, { x: 2.2, y: 0.6, w: 5.6, h: 4.2, fill: { color: "EDF3FF" }, line: { color: "EDF3FF", width: 0 } });
      endSlide.addText("THANK YOU", {
        x: 0.5,
        y: 1.7,
        w: 9,
        h: 0.95,
        fontFace: "Cambria",
        fontSize: 54,
        color: PPT.slate,
        bold: true,
        align: "center",
        valign: "mid"
      });
      endSlide.addText("End of quiz presentation", {
        x: 0.5,
        y: 3.28,
        w: 9,
        h: 0.38,
        fontFace: "Calibri",
        fontSize: 21,
        color: PPT.depedBlue,
        align: "center"
      });
      addFooter(endSlide);

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
    <div
      style={{
        fontFamily: "'Trebuchet MS', 'Gill Sans', 'Segoe UI', sans-serif",
        minHeight: "100vh",
        background: "radial-gradient(circle at top left, #f8fbff 0%, #e7edf6 42%, #dde6f2 100%)",
        color: DEPED.text,
        padding: "24px 14px 40px",
        fontSize: "16px"
      }}
    >
      <style>{`
        @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
        @keyframes pulse { 0%,100% { opacity:1; } 50% { opacity:0.45; } }
      `}</style>

      <div style={{ maxWidth: 960, margin: "0 auto" }}>
        <div style={{ ...neuRaised(26), padding: "18px 24px", marginBottom: 20 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
            <div
              style={{
                ...neuRaised(16),
                width: 48,
                height: 48,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                color: DEPED.blue,
                border: `2px solid ${DEPED.blue}`,
                fontSize: 24
              }}
            >
              📘
            </div>
            <div style={{ flex: 1, minWidth: 220 }}>
              <div style={{ fontSize: 28, fontWeight: 800, color: DEPED.blue, letterSpacing: 0.4 }}>QuizSlide Maker</div>
              <div style={{ fontSize: 12, color: DEPED.muted, letterSpacing: 2.3 }}>DEPED READY · UPLOAD · EXTRACT · PRESENT</div>
            </div>
            {questions.length > 0 && (
              <div
                style={{
                  ...neuInset(18),
                  padding: "10px 16px",
                  fontSize: 14,
                  color: DEPED.red,
                  fontWeight: 700,
                  border: `1px solid ${DEPED.border}`
                }}
              >
                {questions.length} questions extracted
              </div>
            )}
          </div>
        </div>

        <div style={{ ...neuRaised(26), padding: "24px 20px", marginBottom: 20 }}>
          <div style={{ display: "flex", justifyContent: "center", gap: 0, marginBottom: 6, flexWrap: "wrap" }}>
            {["Upload", "Extract", "Preview", "Generate"].map((label, i) => {
              const stageMap = ["upload", "processing", "preview", "generating"];
              const doneIdx = ["upload", "processing", "preview", "generating", "done"].indexOf(stage);
              const isActive = stageMap[i] === stage || (stage === "done" && i === 3);
              const isDone = doneIdx > i;
              return (
                <div key={i} style={{ display: "flex", alignItems: "center", marginBottom: 8 }}>
                  <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 6 }}>
                    <div
                      style={{
                        ...neuRaised(999),
                        width: 38,
                        height: 38,
                        border: `2px solid ${isDone ? DEPED.red : isActive ? DEPED.yellow : DEPED.border}`,
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        fontSize: 14,
                        fontWeight: 800,
                        color: isDone ? DEPED.red : isActive ? DEPED.blue : DEPED.muted
                      }}
                    >
                      {isDone ? "✓" : i + 1}
                    </div>
                    <span style={{ fontSize: 12, color: isActive ? DEPED.blue : isDone ? DEPED.red : DEPED.muted, fontWeight: 700 }}>{label}</span>
                  </div>
                  {i < 3 && (
                    <div
                      style={{
                        width: 60,
                        height: 4,
                        borderRadius: 999,
                        margin: "0 10px 20px",
                        background: isDone
                          ? `linear-gradient(90deg, ${DEPED.red}, ${DEPED.yellow})`
                          : `linear-gradient(90deg, ${DEPED.surfaceAlt}, ${DEPED.surface})`
                      }}
                    />
                  )}
                </div>
              );
            })}
          </div>
        </div>

        {error && (
          <div style={{ ...neuRaised(18), border: `1px solid #e8a3ab`, padding: "14px 16px", marginBottom: 16, color: DEPED.error, fontSize: 13, wordBreak: "break-word" }}>
            <div style={{ fontWeight: 800, marginBottom: 4 }}>Error</div>
            <div>{error}</div>
          </div>
        )}

        {debugLog.length > 0 && (
          <div style={{ ...neuInset(16), border: `1px solid ${DEPED.border}`, padding: "12px 16px", marginBottom: 16, fontSize: 11, fontFamily: "monospace", color: DEPED.muted, maxHeight: 140, overflowY: "auto" }}>
            <div style={{ color: DEPED.blue, fontWeight: 800, marginBottom: 6 }}>DEBUG LOG</div>
            {debugLog.map((l, i) => <div key={i} style={{ marginBottom: 2 }}>› {l}</div>)}
          </div>
        )}

        <div style={{ ...neuRaised(28), padding: "30px 24px" }}>
          {stage === "upload" && (
            <div>
              <div style={{ textAlign: "center", marginBottom: 30 }}>
                <h1 style={{ fontSize: 32, fontWeight: 800, color: DEPED.blue, margin: "0 0 10px" }}>Upload Your Test Paper</h1>
                <p style={{ color: DEPED.muted, fontSize: 16, margin: 0 }}>Supports PDF, DOCX, or TXT test formats. No API key required.</p>
              </div>

              <div
                onDrop={handleDrop}
                onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
                onDragLeave={() => setDragOver(false)}
                onClick={() => fileInputRef.current?.click()}
                style={{
                  ...neuInset(22),
                  border: `2px dashed ${dragOver ? DEPED.yellow : file ? DEPED.red : DEPED.border}`,
                  padding: "48px 28px",
                  textAlign: "center",
                  cursor: "pointer",
                  background: dragOver ? "#f7f1d4" : file ? "#f5e9eb" : DEPED.surface,
                  transition: "all 0.25s",
                  marginBottom: 26
                }}
              >
                <input ref={fileInputRef} type="file" accept=".pdf,.txt,.docx,.doc" style={{ display: "none" }} onChange={(e) => readFile(e.target.files[0])} />
                <div style={{ fontSize: 58, marginBottom: 12 }}>{file ? "📄" : "📁"}</div>
                {file ? (
                  <>
                    <div style={{ fontSize: 20, fontWeight: 700, color: DEPED.blue, marginBottom: 8 }}>{file.name}</div>
                    <div style={{ fontSize: 14, color: DEPED.muted }}>{(file.size / 1024).toFixed(1)} KB · Click to change file</div>
                    {fileData && <div style={{ marginTop: 6, fontSize: 14, color: DEPED.success, fontWeight: 700 }}>File parsed and ready</div>}
                    {!fileData && <div style={{ marginTop: 6, fontSize: 14, color: DEPED.red }}>Parsing file...</div>}
                  </>
                ) : (
                  <>
                    <div style={{ fontSize: 20, fontWeight: 700, color: DEPED.blue, marginBottom: 8 }}>Drop your test paper here</div>
                    <div style={{ fontSize: 15, color: DEPED.muted }}>or click to browse · PDF, DOCX, TXT supported</div>
                  </>
                )}
              </div>

              <div style={{ ...neuInset(16), padding: "16px 18px", marginBottom: 24, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 14 }}>
                <div>
                  <div style={{ fontWeight: 700, color: DEPED.blue, fontSize: 16 }}>Randomize Questions</div>
                  <div style={{ color: DEPED.muted, fontSize: 14, marginTop: 4 }}>Shuffle question order in the presentation</div>
                </div>
                <button
                  onClick={() => setRandomize(!randomize)}
                  aria-label="Toggle randomize"
                  style={{
                    ...neuInset(999),
                    width: 56,
                    height: 30,
                    border: "none",
                    background: randomize ? "#fce9eb" : DEPED.surface,
                    cursor: "pointer",
                    position: "relative"
                  }}
                >
                  <span
                    style={{
                      position: "absolute",
                      top: 4,
                      left: randomize ? 30 : 4,
                      width: 22,
                      height: 22,
                      borderRadius: "50%",
                      background: randomize ? DEPED.red : DEPED.blue,
                      transition: "all 0.3s"
                    }}
                  />
                </button>
              </div>

              <button
                onClick={extractQuestions}
                disabled={!file}
                style={{
                  width: "100%",
                  padding: "16px",
                  borderRadius: 16,
                  border: "none",
                  background: file ? `linear-gradient(135deg, ${DEPED.blue}, #265ec9)` : DEPED.surfaceAlt,
                  color: file ? "#fff" : DEPED.muted,
                  fontSize: 18,
                  fontWeight: 800,
                  cursor: file ? "pointer" : "not-allowed",
                  letterSpacing: 0.4,
                  boxShadow: file ? "8px 8px 18px #b8c4d8, -8px -8px 18px #ffffff" : "none"
                }}
              >
                {status === "Reading file..." ? "Reading file..." : "Extract Questions"}
              </button>
            </div>
          )}

          {stage === "processing" && (
            <div style={{ textAlign: "center", padding: "56px 0" }}>
              <div style={{ fontSize: 62, marginBottom: 18, color: DEPED.blue, animation: "spin 1.5s linear infinite" }}>⚙️</div>
              <h2 style={{ fontSize: 28, color: DEPED.blue, marginBottom: 10, fontWeight: 800 }}>Processing Test Paper</h2>
              <p style={{ color: DEPED.muted, fontSize: 15, animation: "pulse 2s ease-in-out infinite" }}>{status}</p>
              <div style={{ marginTop: 26, display: "flex", justifyContent: "center", gap: 8 }}>
                {[0, 1, 2].map((i) => (
                  <div key={i} style={{ width: 10, height: 10, borderRadius: "50%", background: i === 1 ? DEPED.red : DEPED.yellow, animation: `pulse 1.2s ease-in-out ${i * 0.2}s infinite` }} />
                ))}
              </div>
            </div>
          )}

          {stage === "preview" && (
            <div>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 22, gap: 10, flexWrap: "wrap" }}>
                <div>
                  <h2 style={{ fontSize: 26, fontWeight: 800, color: DEPED.blue, margin: "0 0 6px" }}>{testTitle}</h2>
                  <p style={{ color: DEPED.muted, fontSize: 15, margin: 0 }}>{questions.length} questions extracted. Review before generating.</p>
                </div>
                <button
                  onClick={() => setStage("upload")}
                  style={{ ...neuRaised(12), border: `1px solid ${DEPED.border}`, padding: "9px 14px", color: DEPED.blue, cursor: "pointer", fontSize: 13, fontWeight: 700 }}
                >
                  Back
                </button>
              </div>

              <div style={{ ...neuInset(14), padding: "14px 16px", marginBottom: 18, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 14 }}>
                <div style={{ fontWeight: 700, color: DEPED.blue, fontSize: 14 }}>Randomize in presentation</div>
                <button
                  onClick={() => setRandomize(!randomize)}
                  aria-label="Toggle randomize in preview"
                  style={{
                    ...neuInset(999),
                    width: 56,
                    height: 30,
                    border: "none",
                    background: randomize ? "#fce9eb" : DEPED.surface,
                    cursor: "pointer",
                    position: "relative"
                  }}
                >
                  <span style={{ position: "absolute", top: 4, left: randomize ? 30 : 4, width: 22, height: 22, borderRadius: "50%", background: randomize ? DEPED.red : DEPED.blue, transition: "all 0.3s" }} />
                </button>
              </div>

              <div style={{ maxHeight: 420, overflowY: "auto", marginBottom: 22, display: "flex", flexDirection: "column", gap: 10, paddingRight: 4 }}>
                {questions.map((q, i) => (
                  <div key={i} style={{ ...neuRaised(12), border: `1px solid ${DEPED.border}`, overflow: "hidden" }}>
                    <div onClick={() => setExpandedQ(expandedQ === i ? null : i)} style={{ display: "flex", alignItems: "flex-start", gap: 14, padding: "14px 16px", cursor: "pointer" }}>
                      <div style={{ ...neuInset(999), width: 30, height: 30, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13, fontWeight: 800, color: DEPED.blue, flexShrink: 0, marginTop: 1 }}>{i + 1}</div>
                      <div style={{ flex: 1 }}>
                        <div style={{ color: DEPED.text, fontSize: 15, lineHeight: 1.55 }}>{q.question}</div>
                        {q.choices?.length > 0 && <div style={{ marginTop: 6, fontSize: 13, color: DEPED.muted }}>{q.choices.length} choices</div>}
                      </div>
                      <div style={{ color: DEPED.muted, fontSize: 18, transform: expandedQ === i ? "rotate(180deg)" : "none", transition: "transform 0.2s" }}>▾</div>
                    </div>
                    {expandedQ === i && (
                      <div style={{ padding: "0 16px 14px 58px", borderTop: `1px solid ${DEPED.border}` }}>
                        {q.choices?.length > 0 && (
                          <div style={{ marginTop: 10, display: "flex", flexDirection: "column", gap: 5 }}>
                            {q.choices.map((ch, ci) => (
                              <div key={ci} style={{ ...neuInset(8), fontSize: 14, color: DEPED.text, lineHeight: 1.55, padding: "6px 8px" }}>{ch}</div>
                            ))}
                          </div>
                        )}
                        <div style={{ ...neuInset(8), marginTop: 10, padding: "10px 12px", borderLeft: `3px solid ${DEPED.red}` }}>
                          <div style={{ fontSize: 11, color: DEPED.red, fontWeight: 800, marginBottom: 3 }}>ANSWER</div>
                          <div style={{ fontSize: 15, color: DEPED.text }}>{q.answer}</div>
                        </div>
                        {q.rationale && (
                          <div style={{ ...neuInset(8), marginTop: 8, padding: "10px 12px", borderLeft: `3px solid ${DEPED.yellow}` }}>
                            <div style={{ fontSize: 11, color: DEPED.blue, fontWeight: 800, marginBottom: 3 }}>RATIONALE</div>
                            <div style={{ fontSize: 14, color: DEPED.text, lineHeight: 1.6 }}>{q.rationale}</div>
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                ))}
              </div>

              <button
                onClick={generatePPTX}
                style={{
                  width: "100%",
                  padding: "16px",
                  borderRadius: 16,
                  border: "none",
                  background: `linear-gradient(135deg, ${DEPED.red}, #a40f22)`,
                  color: "#fff",
                  fontSize: 18,
                  fontWeight: 800,
                  cursor: "pointer",
                  letterSpacing: 0.4,
                  boxShadow: "8px 8px 18px #b8c4d8, -8px -8px 18px #ffffff"
                }}
              >
                Generate PowerPoint Presentation
              </button>
            </div>
          )}

          {stage === "generating" && (
            <div style={{ textAlign: "center", padding: "56px 0" }}>
              <div style={{ fontSize: 62, marginBottom: 16, color: DEPED.red }}>🎬</div>
              <h2 style={{ fontSize: 26, color: DEPED.blue, marginBottom: 10, fontWeight: 800 }}>Building Slides</h2>
              <p style={{ color: DEPED.muted, fontSize: 15 }}>{status}</p>
              <div style={{ ...neuInset(10), marginTop: 24, height: 10, overflow: "hidden" }}>
                <div style={{ height: "100%", width: "70%", background: `linear-gradient(90deg, ${DEPED.yellow}, ${DEPED.red})`, borderRadius: 10, animation: "pulse 1.5s ease-in-out infinite" }} />
              </div>
            </div>
          )}

          {stage === "done" && (
            <div style={{ textAlign: "center", padding: "42px 0" }}>
              <div style={{ fontSize: 66, marginBottom: 10 }}>🎉</div>
              <h2 style={{ fontSize: 32, fontWeight: 800, color: DEPED.blue, marginBottom: 10 }}>Presentation Ready</h2>
              <p style={{ color: DEPED.success, fontSize: 17, marginBottom: 30 }}>Your PPTX file has been downloaded automatically.</p>
              <div style={{ ...neuInset(18), padding: "20px", marginBottom: 26, display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 12 }}>
                {[["Slides", `${questions.length * 2 + 2} total`], ["Questions", questions.length], ["Answers", questions.length], ["Order", randomize ? "Randomized" : "Original"]].map(([label, val]) => (
                  <div key={label} style={{ ...neuRaised(12), padding: "14px", textAlign: "center" }}>
                    <div style={{ fontSize: 12, color: DEPED.muted }}>{label}</div>
                    <div style={{ fontSize: 19, fontWeight: 800, color: DEPED.blue, marginTop: 2 }}>{val}</div>
                  </div>
                ))}
              </div>
              <div style={{ display: "flex", gap: 12, justifyContent: "center", flexWrap: "wrap" }}>
                <button
                  onClick={() => { setStage("preview"); }}
                  style={{ ...neuRaised(12), padding: "12px 22px", border: `1px solid ${DEPED.border}`, color: DEPED.blue, fontSize: 15, fontWeight: 700, cursor: "pointer" }}
                >
                  Re-generate
                </button>
                <button
                  onClick={() => { setStage("upload"); setFile(null); setFileData(null); setQuestions([]); setError(""); }}
                  style={{
                    padding: "12px 22px",
                    borderRadius: 12,
                    border: "none",
                    background: `linear-gradient(135deg, ${DEPED.yellow}, #e8b90f)`,
                    color: DEPED.text,
                    fontSize: 15,
                    fontWeight: 800,
                    cursor: "pointer",
                    boxShadow: "8px 8px 18px #b8c4d8, -8px -8px 18px #ffffff"
                  }}
                >
                  New Test Paper
                </button>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
