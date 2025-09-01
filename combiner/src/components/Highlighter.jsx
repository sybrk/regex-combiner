import { useMemo, useState, useRef } from "react";

/**
 * Inline Regex Highlighter (regex101-style)
 * - User types regex and sees highlighted tokens directly inside the input area.
 * - Achieved by layering a transparent textarea over a styled highlighter div.
 * - Tailwind CSS for styling.
 */

const tokenStyle = {
  delimiter: "bg-slate-800 text-slate-100",
  flags: "bg-slate-700 text-slate-100",
  anchor: "bg-fuchsia-100 text-fuchsia-900 ring-1 ring-fuchsia-200",
  "group-open": "bg-emerald-100 text-emerald-900 ring-1 ring-emerald-200",
  "group-close": "bg-emerald-100 text-emerald-900 ring-1 ring-emerald-200",
  lookaround: "bg-teal-100 text-teal-900 ring-1 ring-teal-200",
  "named-group": "bg-emerald-200 text-emerald-950 ring-1 ring-emerald-300",
  "non-capturing": "bg-emerald-200 text-emerald-950 ring-1 ring-emerald-300",
  alternation: "bg-orange-100 text-orange-900 ring-1 ring-orange-200",
  quantifier: "bg-purple-100 text-purple-900 ring-1 ring-purple-200",
  charclass: "bg-sky-100 text-sky-900 ring-1 ring-sky-200",
  "charclass-range": "bg-sky-200 text-sky-950 ring-1 ring-sky-300",
  "charclass-shorthand": "bg-sky-200 text-sky-950 ring-1 ring-sky-300",
  escape: "bg-rose-100 text-rose-900 ring-1 ring-rose-200",
  backreference: "bg-amber-100 text-amber-900 ring-1 ring-amber-200",
  literal: "bg-gray-100 text-gray-900 ring-1 ring-gray-200",
  dot: "bg-gray-200 text-gray-900 ring-1 ring-gray-300",
  "unicode-prop": "bg-indigo-100 text-indigo-900 ring-1 ring-indigo-200",
  "inline-flag": "bg-slate-200 text-slate-950 ring-1 ring-slate-300",
};

const pretty = (s) => s.replaceAll("\n", "\\n");

function tokenizeRegex(input) {
  let body = input;
  let flags = "";
  let delimited = false;

  if (body.startsWith("/") && body.length >= 2) {
    let i = 1;
    let inEscape = false;
    for (; i < body.length; i++) {
      const ch = body[i];
      if (!inEscape && ch === "/") break;
      inEscape = !inEscape && ch === "\\";
    }
    if (i < body.length) {
      flags = body.slice(i + 1);
      body = body.slice(1, i);
      delimited = true;
    }
  }

  const tokens = [];

  const push = (t) => tokens.push(t);

  const readCharClass = (startIdx) => {
    let j = startIdx + 1;
    let escaped = false;
    for (; j < body.length; j++) {
      const ch = body[j];
      if (!escaped && ch === "]") break;
      escaped = !escaped && ch === "\\";
    }
    const raw = body.slice(startIdx, Math.min(j + 1, body.length));
    const inner = raw.slice(1, Math.max(1, raw.length - 1));
    const rangeRe = /(\\.|[^\\])-([^\\]|\\.)/g;
    let m;
    while ((m = rangeRe.exec(inner))) {
      push({
        type: "charclass-range",
        value: m[0],
        start: startIdx + 1 + (m.index || 0),
        end: startIdx + 1 + (m.index || 0) + m[0].length,
        info: `Character range '${m[0]}'`,
      });
    }
    const shorthandRe = /\\[dDsSwW]/g;
    while ((m = shorthandRe.exec(inner))) {
      push({
        type: "charclass-shorthand",
        value: m[0],
        start: startIdx + 1 + (m.index || 0),
        end: startIdx + 1 + (m.index || 0) + m[0].length,
        info: `${m[0]}: shorthand inside class`,
      });
    }
    push({ type: "charclass", value: raw, start: startIdx, end: startIdx + raw.length, info: "Character class" });
    return Math.max(j, startIdx);
  };

  // Right now this tokenizer only detects char classes, extend for other tokens
  for (let i = 0; i < body.length; i++) {
    if (body[i] === "[") {
      i = readCharClass(i);
    } else {
      push({ type: "literal", value: body[i], start: i, end: i + 1, info: "Literal character" });
    }
  }

  return { tokens, body, flags, delimited };
}

export default function RegexHighlighter() {
  const [input, setInput] = useState("");
  const { tokens } = useMemo(() => tokenizeRegex(input), [input]);
  const textareaRef = useRef(null);

  return (
    <div className="p-4 max-w-2xl mx-auto space-y-4">
      <div className="relative w-full font-mono">
        {/* Highlighter Layer */}
        <div className="absolute inset-0 p-2 whitespace-pre-wrap break-all rounded-lg border font-mono text-sm pointer-events-none">
          {tokens.map((t, idx) => (
            <span
              key={idx}
              className={`px-0.5 rounded ${tokenStyle[t.type] || ""}`}
              title={pretty(t.info)}
            >
              {t.value}
            </span>
          ))}
        </div>

        {/* Transparent Textarea Layer */}
        <textarea
          ref={textareaRef}
          className="relative w-full p-2 rounded-lg border font-mono text-sm bg-transparent text-transparent caret-black resize-none"
          rows={3}
          value={input}
          onChange={(e) => setInput(e.target.value)}
          placeholder="Type regex like /foo(\\d+)/i"
        />
      </div>
    </div>
  );
}
