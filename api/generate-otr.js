// ═══════════════════════════════════════════════════════════════════════════
//  IELTS CITY — OTR Tool BCD  |  api/generate-otr.js
//  Rewritten June 2026 against OTR_BCD_Handoff.md (locked spec)
//  Requires: docx ^8.5.0, jszip ^3.10.1
// ═══════════════════════════════════════════════════════════════════════════
'use strict';
const path  = require('path');
const fs    = require('fs');
const JSZip = require('jszip');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageBreak, ImageRun
} = require('docx');

// ── STATIC ASSETS (Vercel /public/) ───────────────────────────────────────
const PUBLIC    = path.join(__dirname, '..', 'public');
const headerImg = fs.readFileSync(path.join(PUBLIC, 'header_banner.png'));
const footerImg = fs.readFileSync(path.join(PUBLIC, 'footer_banner.png'));
const logoImg   = fs.readFileSync(path.join(PUBLIC, 'logo.png'));

// ── COLOURS ───────────────────────────────────────────────────────────────
const NAVY   = '19376E';
const GOLD   = 'FFD241';
const LB     = 'EEF2FB';  // light blue-grey — field fills, col headers (S/W), descriptor (p7)
const LG     = 'F4F5F8';  // light grey — stat cards, inactive CEFR, score cells
const RED    = 'D9534F';  // incorrect answer cells
const BDR    = 'DDE3ED';  // border colour
const STXT   = '595959';  // secondary text (stat labels, Part labels)
const DARK   = '1A1A1A';  // body text, field values
const GREY   = '555555';  // label text p1
const G88    = '888888';  // footer note p7
const BARBG  = 'EEEEEE';  // bar background
const ALTROW = 'FAFAFA';  // even rows accuracy table
const WHITE  = 'FFFFFF';

// ── GEOMETRY ──────────────────────────────────────────────────────────────
const PW = 11906;   // A4 width  (DXA)
const PH = 16838;   // A4 height (DXA)
const CW = 10466;   // content width — same all pages

// Margins  {top, right, bottom, left, header, footer}  all DXA
const M17 = { top:1300, right:720, bottom:900,  left:720, header:720, footer:720 }; // p1 & p7
const M26 = { top:1350, right:720, bottom:720,  left:720, header:720, footer:720 }; // p2–p6

// Feedback column widths (S/W pages) — 22/8/40/30 % of CW
const FC1 = 2303;   // Criterion
const FC2 = 837;    // Score
const FC3 = 4186;   // Descriptor
const FC4 = 3140;   // Notes

// Accuracy table column widths (L/R pages) — 34/12/12/42 % of CW
const AC1 = 3558;   // Question Type
const AC2 = 1256;   // Correct
const AC3 = 1256;   // Total
const AC4 = 4396;   // Accuracy (bar + %)

// Answer map: Part label 806 DXA, each Q cell 966 DXA  (806 + 966×10 = 10466)
const PART_W = 806;
const Q_W    = 966;

// Score table p1: Overall = 2400, skill cols equal-split remainder
const OV_W  = 2400;
const SK_W1 = 2017;  // L
const SK_W2 = 2017;  // R
const SK_W3 = 2016;  // W
const SK_W4 = 2016;  // S  → 2017+2017+2016+2016+2400 = 10466

// ── BORDERS ───────────────────────────────────────────────────────────────
const NIL    = { style: BorderStyle.NIL, size: 0, color: WHITE };
const THIN   = { style: BorderStyle.SINGLE, size: 4, color: BDR };
const THICK  = (c) => ({ style: BorderStyle.SINGLE, size: 12, color: c });
const noB    = { top: NIL, bottom: NIL, left: NIL, right: NIL };
const thB    = { top: THIN, bottom: THIN, left: THIN, right: THIN };
const noBot  = { top: NIL, bottom: NIL, left: NIL, right: NIL };

// ── CELL MARGINS ──────────────────────────────────────────────────────────
const CM  = { top:80,  bottom:80,  left:120, right:120 };
const CML = { top:120, bottom:120, left:200, right:200 };
const CMZ = { top:0,   bottom:0,   left:0,   right:0   };
const CMHZ = { top:60, bottom:60, left:0, right:0 };

// ── CEFR MAPPINGS ─────────────────────────────────────────────────────────
// p1 tile labels (index 0–5 = A1 A2 B1 B2 C1 C2)
const CEFR_CODES  = ['A1','A2','B1','B2','C1','C2'];
const CEFR_NAMES  = ['Beginner','Elementary','Independent User','Independent User','Proficient User','Mastery'];

// p7 pill: overall band → {code, name}
function cefrFromBand(b) {
  if      (b <= 2.5) return { idx:0, code:'A1', name:'Beginner' };
  else if (b <= 3.5) return { idx:1, code:'A2', name:'Elementary' };
  else if (b <= 5.5) return { idx:2, code:'B1', name:'Independent User' };
  else if (b <= 6.5) return { idx:3, code:'B2', name:'Independent User' };
  else if (b <= 7.5) return { idx:4, code:'C1', name:'Proficient User' };
  else               return { idx:5, code:'C2', name:'Mastery' };
}

// ── p7 BAND DESCRIPTORS (Band 2.0–8.0, half-step) ─────────────────────────
// Style guide: second-person, 4–6 sentences, English, no hedging,
// no mention of band number in the text. See Band 6.0 from reference DOCX.
const BAND_DESCRIPTORS = {
  '2.0': 'Your current command of English is at a very early stage of development. You are able to recognise and use a small number of familiar words and short phrases, but communicating beyond the most basic situations remains very challenging. Reading and listening are limited to simple, slow, and repeated input, and written production consists mainly of isolated words or memorised expressions. At this stage, building a foundation of everyday vocabulary and core sentence structures will be your most important priority.',
  '2.5': 'Your English is at an early foundational level, with the ability to handle isolated familiar words and simple phrases. You can follow very short, simple texts and instructions when supported by context or visual aids, and you are beginning to produce basic utterances in familiar situations. Significant effort is still required in all four skills, and communication often breaks down with unfamiliar topics or speakers. Consistent practice with everyday vocabulary and basic grammar patterns will form the core of your development at this stage.',
  '3.0': 'Your English is at an elementary level, allowing you to handle simple, predictable interactions in familiar situations. You can follow straightforward texts and short conversations when the topic is known and the language is slow and clear. Written production is limited to simple sentences on familiar topics, with frequent errors that sometimes affect meaning. Prioritising grammar accuracy, expanding core vocabulary, and practising simple connected writing will help you build the foundation needed to progress.',
  '3.5': 'Your English reflects a developing elementary foundation, and you are beginning to handle simple conversations and short written tasks in familiar contexts. You can understand the main point of slow, clear speech and simple written texts when the topic is predictable. In writing and speaking, you are producing basic sentences, though errors are frequent and sometimes disrupt communication. Focus on building grammatical accuracy in simple structures, expanding your working vocabulary, and practising producing connected sentences in both writing and speech.',
  '4.0': 'Your English reflects a basic working ability across the four skills, and you are able to communicate in familiar, straightforward situations when support is available. You can follow the main ideas of simple texts and slow, clear speech, and your writing and speaking demonstrate effort to produce organised responses, though errors in grammar and vocabulary are still frequent and affect clarity. Developing more control over basic grammar structures, broadening your range of vocabulary, and practising sustained production in both written and spoken tasks will be your priorities at this stage.',
  '4.5': 'Your English shows a working command across familiar situations, and you are able to convey basic meaning in both writing and speaking, though with noticeable limitations. You follow simple to moderately complex texts and conversations when the topic is predictable, and your responses demonstrate some organisation and coherence. Grammatical errors remain frequent, and your vocabulary range limits your ability to express precise meaning. Strengthening grammatical accuracy, broadening active vocabulary, and practising longer, more developed responses in both skills will support your progression toward greater independence.',
  '5.0': 'Your English reflects a modest level of proficiency, and you are able to handle the main demands of everyday communication in familiar contexts. You follow the gist of moderately complex texts and conversations, and your written and spoken responses show some organisation, though coherence and development remain inconsistent. Grammar and vocabulary errors are still noticeable and occasionally disrupt communication. Focus on reducing systematic grammar errors, expanding your vocabulary range to include less common expressions, and developing your ability to sustain a clear argument or description across a full response.',
  '5.5': 'Your English shows a developing level of competence, and you are able to engage with a range of familiar and moderately demanding tasks across all four skills. You can follow the main ideas of complex input when the topic is reasonably accessible, and your written and spoken production is generally understandable, though coherence, accuracy, and range are still areas for development. To move to the next level, focus on producing more complex and accurate grammatical structures, using cohesive devices more naturally in writing, and extending the depth and specificity of your spoken responses.',
  '6.0': 'Your performance reflects a solid upper-intermediate level of English proficiency. You are able to understand the main ideas of complex text on both concrete and abstract topics, and you can interact with a degree of fluency and spontaneity that makes regular interaction with native speakers quite possible. Your writing demonstrates a clear ability to organise ideas and present information coherently, though you would benefit from expanding your use of complex grammatical structures and developing arguments with greater depth and supporting detail. In Speaking, your delivery is generally clear and confident, and you manage to maintain communication effectively across a range of topics. To move toward the next band level, focus on reducing hesitation in extended responses, using a wider range of cohesive devices accurately, and developing the precision of your lexical choices in both written and spoken tasks.',
  '6.5': 'Your English reflects a strong upper-intermediate level, and you demonstrate an ability to engage confidently with complex language in most everyday and academic contexts. You follow extended discourse and nuanced argument with reasonable ease, and your written and spoken production is generally well-organised and effective, though occasional inaccuracies and limited range in grammar or vocabulary sometimes reduce your impact. To consolidate your progress, work on extending the precision of your language choices, producing a wider variety of complex grammatical structures with consistent accuracy, and developing the ability to sustain detailed, well-supported arguments in both written and spoken tasks.',
  '7.0': 'Your English reflects a strong level of proficiency, and you are able to engage successfully with most academic and professional language demands. You follow extended discourse across a wide range of topics, and your writing demonstrates clear organisation, coherent argument development, and a generally good range of vocabulary and grammar. In speaking, you communicate effectively and with confidence, though occasional inaccuracies or lapses in fluency still occur in more demanding situations. To reach the next level, direct your attention to eliminating systematic errors in complex grammar, refining the precision and naturalness of your vocabulary choices, and developing fuller and more convincing support for the arguments you make in writing and speech.',
  '7.5': 'Your English is at an advanced level, and you are able to operate effectively across a wide range of demanding language contexts. You follow complex and abstract ideas with confidence, and your writing and speaking demonstrate sophisticated organisation, a strong range of vocabulary, and generally accurate grammar. Errors, when they occur, are minor and rarely affect communication. To move toward the highest band levels, focus on the precision of your language choices, the naturalness of your idiomatic expression, and your ability to construct and sustain highly nuanced arguments with full control of register and style.',
  '8.0': 'Your English reflects a very high level of proficiency across all four skills. You can handle complex, abstract, and unfamiliar language with ease and precision, and your writing demonstrates sophisticated organisation, a wide vocabulary range, and near-faultless grammatical accuracy. In speaking, you express yourself fluently and spontaneously, with only very occasional minor slips that do not affect comprehension. At this level, your focus should be on achieving complete mastery of register, idiom, and stylistic flexibility — the qualities that distinguish a very good user from an expert one.'
};

function getBandDescriptor(band) {
  const key = parseFloat(band).toFixed(1);
  return BAND_DESCRIPTORS[key] || BAND_DESCRIPTORS['6.0'];
}

// ── SUB-CRITERION LABELS ──────────────────────────────────────────────────
// Verbatim from reference calibration document. Do not alter.
// S tab sub-criteria (in order matching payload subCriteria arrays)
const S_FC  = ['Độ trôi chảy', 'Ngôn ngữ nối', 'Độ mạch lạc'];
const S_LR  = ['Phổ từ vựng', 'Độ chính xác', 'Paraphrase'];
const S_GRA = ['Phổ ngữ pháp', 'Độ chính xác'];
const S_PRN = ['Phổ đặc điểm phát âm', 'Khả năng kiểm soát phát âm', 'Độ dễ hiểu'];

// W1 tab sub-criteria
const W1_TA  = ['Đáp ứng yêu cầu đề bài', 'Góc nhìn tổng quan', 'Trình bày thông tin & dữ liệu'];
const W1_CC  = ['Tổ chức thông tin', 'Liên kết câu & đoạn', 'Chia đoạn'];
const W1_LR  = ['Phổ từ vựng', 'Độ chính xác'];
const W1_GRA = ['Phổ ngữ pháp', 'Độ chính xác'];

// W2 tab sub-criteria
const W2_TA  = ['Đáp ứng yêu cầu đề bài', 'Trình bày quan điểm', 'Trình bày & phát triển luận điểm'];
const W2_CC  = ['Tổ chức bài viết', 'Liên kết câu & đoạn', 'Chia đoạn'];
const W2_LR  = ['Phổ từ vựng', 'Độ chính xác'];
const W2_GRA = ['Phổ ngữ pháp', 'Độ chính xác'];

// Map criterion name → sub-criterion label arrays
const S_LABELS  = { 'FLUENCY & COHERENCE': S_FC, 'LEXICAL RESOURCE': S_LR, 'GRAMMATICAL RANGE & ACCURACY': S_GRA, 'PRONUNCIATION': S_PRN };
const W1_LABELS = { 'TASK ACHIEVEMENT': W1_TA, 'COHERENCE & COHESION': W1_CC, 'COHERENCE AND COHESION': W1_CC, 'LEXICAL RESOURCE': W1_LR, 'GRAMMATICAL RANGE & ACCURACY': W1_GRA };
const W2_LABELS = { 'TASK ACHIEVEMENT': W2_TA, 'TASK RESPONSE': W2_TA, 'COHERENCE & COHESION': W2_CC, 'COHERENCE AND COHESION': W2_CC, 'LEXICAL RESOURCE': W2_LR, 'GRAMMATICAL RANGE & ACCURACY': W2_GRA };

function getSubLabels(criterionName, labelMap) {
  const key = (criterionName || '').toUpperCase();
  if (labelMap[key]) return labelMap[key];
  const found = Object.keys(labelMap).find(k => key.includes(k) || k.includes(key));
  return found ? labelMap[found] : [];
}

// ── SUB-CRITERION DESCRIPTOR LOOKUP ───────────────────────────────────────
// Key: "subcriterion_label:band_score" e.g. "Độ trôi chảy:6.0"
// Populated from reference DOCX sample. Missing entries → '—' placeholder.
// Full table will be supplied from calibration document separately.
const DESCRIPTORS = {
  // SPEAKING — FLUENCY & COHERENCE
  'Độ trôi chảy:6.0':   'Bạn không thể trả lời mà không có những khoảng dừng rõ ràng; thường phải nói chậm, lặp lại hoặc tự sửa để tìm được ngôn ngữ cần dùng.',
  'Ngôn ngữ nối:6.5':   'Bạn sử dụng phổ từ nối tương đối rộng nhưng đôi chỗ còn chưa phù hợp với nghĩa cần diễn đạt.',
  'Độ mạch lạc:6.0':    'Bạn phát triển được các ý đơn giản khá rõ ràng, nhưng gặp khó khăn khi mở rộng hoặc đào sâu ý với các chủ đề phức tạp hơn.',
  // SPEAKING — LEXICAL RESOURCE
  'Phổ từ vựng:6.5':    'Bạn có đủ từ vựng để thảo luận về nhiều chủ đề khác nhau, kể cả các chủ đề ít quen thuộc.',
  'Độ chính xác:6.0':   'Bạn còn dùng từ chưa chính xác khá thường xuyên nhưng ý muốn truyền đạt vẫn rõ ràng với người nghe.',
  'Paraphrase:7.0':      'Bạn thường paraphrase khá thành công khi không tìm được từ chính xác; chỉ đôi khi cách diễn đạt thay thế còn chưa thực sự tự nhiên.',
  // SPEAKING — GRAMMATICAL RANGE & ACCURACY
  'Phổ ngữ pháp:6.0':   'Bạn sử dụng hỗn hợp các cấu trúc đơn giản và phức tạp, tuy nhiên các cấu trúc phức vẫn còn chiếm tỉ lệ ít hơn.',
  // Note: 'Độ chính xác:6.0' already registered above — same key works for GRA too
  // SPEAKING — PRONUNCIATION
  'Phổ đặc điểm phát âm:7.0':    'Bạn thể hiện phổ rộng các đặc điểm phát âm tiếng Anh một cách ổn định, bao gồm kết nối âm và ngữ điệu tự nhiên.',
  'Khả năng kiểm soát phát âm:7.0': 'Bạn kiểm soát phát âm khá hiệu quả trong hầu hết các tình huống, tuy nhiên chưa duy trì được xuyên suốt khi nói nhanh hoặc gặp từ khó.',
  'Độ dễ hiểu:7.0':     'Phát âm của bạn nhìn chung dễ hiểu xuyên suốt; chỉ đôi khi một số từ hoặc âm bị phát âm sai làm gián đoạn người nghe.',
  // WRITING TASK 1 — TASK ACHIEVEMENT
  'Đáp ứng yêu cầu đề bài:6.0':  'Bạn đáp ứng được các yêu cầu đề bài — mô tả đúng dạng và bao quát được dữ liệu được yêu cầu.',
  'Góc nhìn tổng quan:6.0':       'Bạn đưa ra được góc nhìn tổng quan với các xu hướng và thông tin chính được lựa chọn phù hợp.',
  'Trình bày thông tin & dữ liệu:6.0': 'Bạn trình bày và nhấn mạnh được các thông tin chính; các chi tiết bổ trợ nhìn chung phù hợp, dù đôi chỗ còn chưa chính xác hoặc chưa liên quan.',
  // WRITING TASK 1 — COHERENCE & COHESION
  'Tổ chức thông tin:6.0':        'Bạn sắp xếp thông tin mạch lạc; người đọc có thể theo dõi được cách bạn tổ chức và trình bày dữ liệu.',
  'Liên kết câu & đoạn:5.5':      'Bạn sử dụng từ nối khá hiệu quả nhưng đôi khi còn máy móc hoặc chưa phù hợp; đại từ và từ thay thế chưa được dùng nhất quán.',
  'Liên kết câu & đoạn:6.0':      'Bạn sử dụng từ nối khá hiệu quả nhưng đôi khi còn máy móc hoặc chưa phù hợp; đại từ và từ thay thế chưa được dùng nhất quán.',
  'Chia đoạn:6.0':                'Bạn chia đoạn tương đối hợp lý nhưng đôi khi một đoạn chứa nhiều hơn một ý chính, hoặc chất lượng chia đoạn chưa đồng đều giữa các phần của bài.',
  // WRITING TASK 1 — LEXICAL RESOURCE
  'Phổ từ vựng:6.5':              'Bạn sử dụng phổ từ vựng đủ rộng để đáp ứng yêu cầu đề và bắt đầu sử dụng một số từ ít thông dụng.',
  'Độ chính xác:6.5':             'Sử dụng phổ từ vựng rộng và phong phú hơn, bao gồm cách diễn đạt tự nhiên.',
  // WRITING TASK 1 — GRA
  'Phổ ngữ pháp:6.0':             'Sử dụng phổ rộng và đa dạng các cấu trúc phức, bao gồm nhiều loại khác nhau.',
  // 'Độ chính xác:6.0' already registered above
  // WRITING TASK 2 — TASK RESPONSE
  'Đáp ứng yêu cầu đề bài:6.5':  'Bạn đã đề cập đến tất cả các phần của đề, nhưng một số phần được phát triển đầy đủ hơn các phần còn lại.',
  'Trình bày quan điểm:6.5':      'Bạn trình bày được quan điểm liên quan tuy nhiên các kết luận có thể còn chưa rõ ràng hoặc lặp đi lặp lại.',
  'Trình bày & phát triển luận điểm:6.0': 'Bạn trình bày được các ý chính nhưng dừng lại ở mức nêu ý — chưa có đủ giải thích, lý lẽ hoặc ví dụ cụ thể để làm rõ và thuyết phục người đọc.',
  // WRITING TASK 2 — COHERENCE & COHESION
  'Tổ chức bài viết:7.0':         'Bạn sắp xếp thông tin và ý tưởng mạch lạc; người đọc có thể theo dõi được hướng phát triển của bài.',
  'Chia đoạn:6.5':                'Bạn chia đoạn tương đối hợp lý nhưng đôi khi một đoạn chứa nhiều hơn một ý chính, hoặc chất lượng chia đoạn chưa đồng đều giữa các phần của bài.',
  // WRITING TASK 2 — GRA
  'Phổ ngữ pháp:7.0':             'Sử dụng phổ rộng và đa dạng các cấu trúc phức, bao gồm nhiều loại khác nhau.',
  'Độ chính xác:7.0':             'Đa số các câu không có lỗi; chỉ còn rất ít lỗi nhỏ không ảnh hưởng đến giao tiếp.'
};

function getDescriptor(subName, score) {
  const key = subName + ':' + parseFloat(score).toFixed(1);
  return DESCRIPTORS[key] || '—'; // em-dash for unresolved entries
}

// ── BASIC HELPERS ─────────────────────────────────────────────────────────
const t   = (text, o = {}) => new TextRun({ text: String(text), font: 'Arial', size: 18, color: DARK, ...o });
const p   = (children, o = {}) => new Paragraph({ children, spacing: { before: 0, after: 0 }, ...o });
const ep  = (sz = 6)  => p([t('', { size: sz })]);
const pb  = ()        => p([new PageBreak()]);

function row(cells) { return new TableRow({ children: cells }); }
function cell(children, opts = {}) {
  const ch = Array.isArray(children) ? children : [p([t(String(children))])];
  return new TableCell({ borders: thB, margins: CM, verticalAlign: VerticalAlign.CENTER, ...opts, children: ch });
}
function navyCell(children, opts = {}) {
  return cell(children, { shading: { fill: NAVY, type: ShadingType.CLEAR }, borders: noB, ...opts });
}
function lbCell(children, opts = {}) {
  return cell(children, { shading: { fill: LB, type: ShadingType.CLEAR }, ...opts });
}

// One-row table spanning CW
function oneRowTable(cells, colWidths) {
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [row(cells)],
    borders: noB
  });
}

// ── HEADER / FOOTER ───────────────────────────────────────────────────────
function makeHeader() {
  return new Header({ children: [p([new ImageRun({ data: headerImg, transformation: { width: 697, height: 33 }, type: 'png' })])] });
}
function makeFooter() {
  return new Footer({ children: [p([new ImageRun({ data: footerImg, transformation: { width: 571, height: 34 }, type: 'png' })], { alignment: AlignmentType.CENTER })] });
}

// ── SECTION TITLE (p1) ────────────────────────────────────────────────────
// 13pt bold navy, 6pt navy bottom border
function sectionTitle(text) {
  return p([t(text, { bold: true, color: NAVY, size: 26 })], {
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY } }
  });
}

// ── SKILL BANNER (pp 4–6) ────────────────────────────────────────────────
// Left ~75% navy, right ~25% gold
function skillBanner(title, bandScore) {
  const lw = Math.round(CW * 0.75);
  const rw = CW - lw;
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [lw, rw],
    rows: [row([
      navyCell([p([t(title, { bold: true, color: WHITE, size: 32 })])], { width: { size: lw, type: WidthType.DXA }, margins: CML }),
      cell([p([t('Band Score: ' + bandScore, { bold: true, color: NAVY, size: 26 })], { alignment: AlignmentType.CENTER })],
        { shading: { fill: GOLD, type: ShadingType.CLEAR }, borders: noB, width: { size: rw, type: WidthType.DXA }, margins: CML })
    ])]
  });
}

// ── FEEDBACK COLUMN HEADER ROW ───────────────────────────────────────────
function feedbackColHeader() {
  const cols = ['Criterion','Score','Descriptor','Notes'];
  const ws   = [FC1, FC2, FC3, FC4];
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: ws,
    rows: [row(cols.map((lbl, i) => cell(
      [p([t(lbl, { bold: true, color: NAVY, size: 22 })], { alignment: i===1 ? AlignmentType.CENTER : AlignmentType.LEFT })],
      { shading: { fill: LB, type: ShadingType.CLEAR }, width: { size: ws[i], type: WidthType.DXA }, margins: CM }
    )))]
  });
}

// ── CRITERION TABLE (one per criterion) ──────────────────────────────────
function criterionTable(name, band, subCriteria, labelMap) {
  const labels = getSubLabels(name, labelMap);
  const bannerRow = row([
    new TableCell({
      columnSpan: 4,
      width: { size: CW, type: WidthType.DXA },
      shading: { fill: NAVY, type: ShadingType.CLEAR },
      borders: noB,
      margins: CM,
      verticalAlign: VerticalAlign.CENTER,
      children: [p([
        t(name.toUpperCase(), { bold: true, color: WHITE, size: 24 }),
        t('    ' + band, { bold: true, color: GOLD, size: 24 })
      ])]
    })
  ]);
  const subRows = (subCriteria || []).map((sc, i) => {
    const lbl  = labels[i] || ('Sub-criterion ' + (i + 1));
    const desc = getDescriptor(lbl, sc.score);
    return row([
      cell([p([t(lbl, { bold: true, color: DARK, size: 18 })])],
        { width: { size: FC1, type: WidthType.DXA }, margins: CM }),
      cell([p([t(String(sc.score), { color: DARK, size: 18 })], { alignment: AlignmentType.CENTER })],
        { shading: { fill: LG, type: ShadingType.CLEAR }, width: { size: FC2, type: WidthType.DXA }, margins: CM }),
      cell([p([t(desc, { size: 18, color: DARK })])],
        { width: { size: FC3, type: WidthType.DXA }, margins: CM }),
      cell([p([t(sc.notes || '', { size: 18, color: DARK, italics: !!sc.notes })])],
        { width: { size: FC4, type: WidthType.DXA }, margins: CM })
    ]);
  });
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [FC1, FC2, FC3, FC4],
    rows: [bannerRow, ...subRows]
  });
}

// ── STAT CARDS (L/R pages) ───────────────────────────────────────────────
function statCards(totalCorrect, bandScore) {
  const hw = Math.floor(CW / 2);
  function makeCard(label, value) {
    return cell([
      p([t(label, { bold: true, color: STXT, size: 17 })]),
      ep(6),
      p([t(String(value), { bold: true, color: NAVY, size: 28 })])
    ], {
      shading: { fill: LG, type: ShadingType.CLEAR },
      width: { size: hw, type: WidthType.DXA },
      margins: CML
    });
  }
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [hw, CW - hw],
    rows: [row([
      makeCard('TOTAL CORRECT', totalCorrect + ' / 40'),
      makeCard('BAND SCORE', bandScore)
    ])]
  });
}

// ── ACCURACY BAR ─────────────────────────────────────────────────────────
// Nested 2-cell bar inside the Accuracy column cell
function accuracyBarCell(correct, total) {
  const pct     = total > 0 ? Math.round((correct / total) * 100) : 0;
  const BAR_W   = Math.round(AC4 * 0.60);   // 60% of accuracy col for the bar table
  const navW    = Math.round(BAR_W * pct / 100);
  const emptyW  = BAR_W - navW;
  const barTable = new Table({
    width: { size: BAR_W, type: WidthType.DXA },
    columnWidths: [navW || 1, emptyW || BAR_W - 1],
    rows: [row([
      new TableCell({ width: { size: navW || 1, type: WidthType.DXA }, shading: { fill: NAVY, type: ShadingType.CLEAR }, borders: noB, margins: CMZ, children: [p([t('', { size: 10 })])] }),
      new TableCell({ width: { size: emptyW || BAR_W - 1, type: WidthType.DXA }, shading: { fill: BARBG, type: ShadingType.CLEAR }, borders: noB, margins: CMZ, children: [p([t('', { size: 10 })])] })
    ])]
  });
  return new TableCell({
    width: { size: AC4, type: WidthType.DXA },
    borders: thB,
    margins: CM,
    verticalAlign: VerticalAlign.CENTER,
    children: [
      ep(4),
      barTable,
      ep(4),
      p([t(pct + '%', { size: 18, color: STXT })], { alignment: AlignmentType.CENTER })
    ]
  });
}

// ── ACCURACY TABLE (L/R pages) ───────────────────────────────────────────
function accuracyTable(byType) {
  const colHeaders = ['Question Type','Correct','Total','Accuracy'];
  const cws = [AC1, AC2, AC3, AC4];
  const headerRow = row(colHeaders.map((lbl, i) =>
    new TableCell({
      width: { size: cws[i], type: WidthType.DXA },
      margins: CM,
      borders: { top: NIL, bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY }, left: NIL, right: NIL },
      children: [p([t(lbl, { bold: true, color: NAVY, size: 22 })], { alignment: i>0 ? AlignmentType.CENTER : AlignmentType.LEFT })]
    })
  ));
  const dataRows = (byType || []).map((bt, idx) => {
    const fill = idx % 2 === 1 ? ALTROW : WHITE;
    return new TableRow({
      children: [
        new TableCell({ width: { size: AC1, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, borders: { top: NIL, bottom: { style: BorderStyle.SINGLE, size: 4, color: 'E5E5E5' }, left: NIL, right: NIL }, margins: CM, children: [p([t(bt.type || '', { size: 18, color: DARK })])] }),
        new TableCell({ width: { size: AC2, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, borders: { top: NIL, bottom: { style: BorderStyle.SINGLE, size: 4, color: 'E5E5E5' }, left: NIL, right: NIL }, margins: CM, children: [p([t(String(bt.correct), { size: 18, color: DARK })], { alignment: AlignmentType.CENTER })] }),
        new TableCell({ width: { size: AC3, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR }, borders: { top: NIL, bottom: { style: BorderStyle.SINGLE, size: 4, color: 'E5E5E5' }, left: NIL, right: NIL }, margins: CM, children: [p([t(String(bt.total), { size: 18, color: DARK })], { alignment: AlignmentType.CENTER })] }),
        accuracyBarCell(bt.correct, bt.total)
      ]
    });
  });
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: cws,
    rows: [headerRow, ...dataRows]
  });
}

// ── ANSWER MAP (L/R pages) ───────────────────────────────────────────────
function answerMap(answerMapArr) {
  const wrong = new Set();
  (answerMapArr || []).forEach((ok, i) => { if (!ok) wrong.add(i + 1); });
  const mapRows = [1,2,3,4].map(part => {
    const start = (part - 1) * 10 + 1;
    const qCells = [];
    for (let q = start; q < start + 10; q++) {
      const isWrong = wrong.has(q);
      qCells.push(new TableCell({
        width: { size: Q_W, type: WidthType.DXA },
        shading: { fill: isWrong ? RED : NAVY, type: ShadingType.CLEAR },
        borders: { top: { style: BorderStyle.SINGLE, size: 8, color: WHITE }, bottom: { style: BorderStyle.SINGLE, size: 8, color: WHITE }, left: { style: BorderStyle.SINGLE, size: 8, color: WHITE }, right: { style: BorderStyle.SINGLE, size: 8, color: WHITE } },
        margins: CMZ,
        verticalAlign: VerticalAlign.CENTER,
        children: [p([t(String(q), { bold: true, color: WHITE, size: 18 })], { alignment: AlignmentType.CENTER })]
      }));
    }
    return new TableRow({
      cantSplit: true,
      children: [
        new TableCell({
          width: { size: PART_W, type: WidthType.DXA },
          borders: noB,
          margins: { top: 80, bottom: 80, left: 0, right: 120 },
          verticalAlign: VerticalAlign.CENTER,
          children: [p([t('Part ' + part, { bold: true, color: STXT, size: 18 })], { alignment: AlignmentType.RIGHT })]
        }),
        ...qCells
      ]
    });
  });
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [PART_W, ...Array(10).fill(Q_W)],
    rows: mapRows
  });
}

// Sub-header for L/R sections
function subHeader(text) {
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    rows: [row([new TableCell({
      width: { size: CW, type: WidthType.DXA },
      shading: { fill: LG, type: ShadingType.CLEAR },
      borders: { top: NIL, bottom: NIL, left: { style: BorderStyle.SINGLE, size: 12, color: NAVY }, right: NIL },
      margins: { top: 100, bottom: 100, left: 200, right: 200 },
      children: [p([t(text, { bold: true, color: '222222', size: 26 })])]
    })])]
  });
}

// ── PAGE 1 ─────────────────────────────────────────────────────────────────
function buildPage1(data) {
  const { candidate, admin, scores } = data;
  const overall = parseFloat(scores.overall || 0);
  const cefr    = cefrFromBand(overall);

  // 1. Logo + title header band
  const logoW = 1800, logoH = 1060;  // px from reference DOCX (168×99 px original, scaled)
  const logoW_px = 168, logoH_px = 99;
  const badgeW = 1800;
  const logoTitleW = CW - badgeW;

  const headerBand = new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [logoTitleW, badgeW],
    rows: [row([
      new TableCell({
        width: { size: logoTitleW, type: WidthType.DXA },
        borders: noB, margins: CMZ, verticalAlign: VerticalAlign.TOP,
        children: [
          p([new ImageRun({ data: logoImg, transformation: { width: logoW_px, height: logoH_px }, type: 'png' })],
            { spacing: { before: 0, after: 80 } }),
          p([t('OFFICIAL TEST REPORT', { bold: true, color: NAVY, size: 44 })],
            { spacing: { before: 80, after: 0 } })
        ]
      }),
      new TableCell({
        width: { size: badgeW, type: WidthType.DXA },
        shading: { fill: NAVY, type: ShadingType.CLEAR },
        borders: noB, margins: CML, verticalAlign: VerticalAlign.CENTER,
        children: [
          p([t('ACADEMIC', { bold: true, color: GOLD, size: 44 })], { alignment: AlignmentType.CENTER }),
          p([t('IELTS CITY', { bold: true, color: WHITE, size: 22 })], { alignment: AlignmentType.CENTER })
        ]
      })
    ])]
  });

  // 2. Disclaimer block — full-width light-blue
  const disclaimer = new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    rows: [row([new TableCell({
      width: { size: CW, type: WidthType.DXA },
      shading: { fill: LB, type: ShadingType.CLEAR },
      borders: noB, margins: { top: 100, bottom: 100, left: 200, right: 200 },
      children: [p([
        t('This report is issued by IELTS CITY as an internal assessment record. It is not an official IELTS certificate issued by the British Council, IDP, or Cambridge Assessment English. Scores are indicative and based on the IELTS CITY marking system.', { size: 17, color: '333333' })
      ])]
    })])]
  });

  // 3. Meta row — Class Code + Date
  const lblW = 2000, valW = (CW - lblW * 2) / 2 | 0;
  const metaRow = new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [lblW, valW, lblW, CW - lblW * 2 - valW],
    rows: [row([
      cell([p([t('Class Code', { size: 18, color: GREY })])], { borders: noB, margins: CM }),
      lbCell([p([t(admin.classCode || '', { bold: true, size: 18 })])], { borders: noB, margins: CM }),
      cell([p([t('Date', { size: 18, color: GREY })])], { borders: noB, margins: CM }),
      lbCell([p([t(admin.testDate || '', { bold: true, size: 18 })])], { borders: noB, margins: CM })
    ])]
  });

  // 4. CANDIDATE DETAILS section title
  const candTitle = sectionTitle('CANDIDATE DETAILS');

  // 5. Name + Photo layout
  // Left: 2-row names table. Right: 1800 DXA photo placeholder column
  const nameW = CW - 1800 - 240;  // photo=1800, gap=240
  const nameLabelW = 2700, nameValW = nameW - nameLabelW;
  const nameTable = new Table({
    width: { size: nameW, type: WidthType.DXA },
    columnWidths: [nameLabelW, nameValW],
    rows: [
      row([
        cell([p([t('Family Name', { size: 20, color: GREY })])], { borders: noB, margins: { top: 80, bottom: 40, left: 0, right: 120 } }),
        lbCell([p([t(candidate.familyName || '', { bold: true, size: 20 })])], { borders: { top: NIL, bottom: NIL, left: NIL, right: NIL }, margins: CM })
      ]),
      row([
        cell([p([t('', { size: 6 })])], { borders: noB, margins: CMZ, children: [ep(6)] }),
        cell([p([t('', { size: 6 })])], { borders: noB, margins: CMZ, children: [ep(6)] })
      ]),
      row([
        cell([p([t('First Name', { size: 20, color: GREY })])], { borders: noB, margins: { top: 40, bottom: 80, left: 0, right: 120 } }),
        lbCell([p([t(candidate.firstName || '', { bold: true, size: 20 })])], { borders: { top: NIL, bottom: NIL, left: NIL, right: NIL }, margins: CM })
      ])
    ]
  });

  // Photo placeholder (text, will be replaced by Fix 3 floating shape)
  const photoPlaceholder = new Table({
    width: { size: 1800, type: WidthType.DXA },
    columnWidths: [1800],
    rows: [row([new TableCell({
      width: { size: 1800, type: WidthType.DXA },
      borders: { top: THIN, bottom: THIN, left: THIN, right: THIN },
      margins: CMZ, verticalAlign: VerticalAlign.CENTER,
      children: [p([t('Photo', { size: 18, color: STXT })], { alignment: AlignmentType.CENTER })]
    })])]
  });

  const namePhotoRow = new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [nameW, 240, 1800],
    rows: [row([
      new TableCell({ width: { size: nameW, type: WidthType.DXA }, borders: noB, margins: CMZ, children: [nameTable] }),
      new TableCell({ width: { size: 240, type: WidthType.DXA }, borders: noB, margins: CMZ, children: [ep()] }),
      new TableCell({ width: { size: 1800, type: WidthType.DXA }, borders: noB, margins: CMZ, children: [photoPlaceholder] })
    ])]
  });

  // Divider
  const divider = p([], { border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BDR } }, spacing: { before: 80, after: 80 } });

  // 7. DOB / Sex / Course Type
  const dob = candidate.dateOfBirth || '';
  const sex = candidate.sex || '';
  const courseType = candidate.courseType || '';
  const sp = 500;
  const f1W = 2000, v1W = 1400;  // DOB label + value
  const f2W = 600,  v2W = 600;   // Sex label + value
  const f3W = 2000, v3W = CW - f1W - v1W - sp - f2W - v2W - sp - f3W;
  const dobRow = new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [f1W, v1W, sp, f2W, v2W, sp, f3W, v3W],
    rows: [row([
      cell([p([t('Date of Birth', { size: 18, color: GREY })])], { borders: noB, margins: CM }),
      lbCell([p([t(dob, { bold: true, size: 18 })])], { borders: noB, margins: CM }),
      new TableCell({ width: { size: sp, type: WidthType.DXA }, borders: noB, margins: CMZ, children: [ep()] }),
      cell([p([t('Sex', { size: 18, color: GREY })])], { borders: noB, margins: CM }),
      lbCell([p([t(sex, { bold: true, size: 18 })])], { borders: noB, margins: CM }),
      new TableCell({ width: { size: sp, type: WidthType.DXA }, borders: noB, margins: CMZ, children: [ep()] }),
      cell([p([t('Course Type', { size: 18, color: GREY })])], { borders: noB, margins: CM }),
      lbCell([p([t(courseType, { bold: true, size: 18 })])], { borders: noB, margins: CM })
    ])]
  });

  // 8. Country of Nationality + First Language
  const clW = 2700, cvW = CW - clW;
  const countryTable = new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [clW, cvW],
    rows: [
      row([
        cell([p([t('Country of Nationality', { size: 18, color: GREY })])], { borders: noB, margins: { top: 80, bottom: 40, left: 0, right: 120 } }),
        lbCell([p([t(candidate.nationality || '', { bold: true, size: 18 })])], { borders: { top: NIL, bottom: NIL, left: NIL, right: NIL }, margins: CM })
      ]),
      row([
        new TableCell({ width: { size: clW, type: WidthType.DXA }, borders: noB, margins: CMZ, children: [ep(6)] }),
        new TableCell({ width: { size: cvW, type: WidthType.DXA }, borders: noB, margins: CMZ, children: [ep(6)] })
      ]),
      row([
        cell([p([t('First Language', { size: 18, color: GREY })])], { borders: noB, margins: { top: 40, bottom: 80, left: 0, right: 120 } }),
        lbCell([p([t(candidate.firstLanguage || '', { bold: true, size: 18 })])], { borders: { top: NIL, bottom: NIL, left: NIL, right: NIL }, margins: CM })
      ])
    ]
  });

  // 10. TEST RESULT section title
  const resultTitle = sectionTitle('TEST RESULT');

  // 11. Score table — L R W S | Overall Band
  const scoreColW = [SK_W1, SK_W2, SK_W3, SK_W4, OV_W];
  const skillLabels = ['Listening', 'Reading', 'Writing', 'Speaking'];
  const skillScores = [scores.listening, scores.reading, scores.writing, scores.speaking];

  const scoreHeaderRow = new TableRow({ children: [
    ...skillLabels.map((lbl, i) => new TableCell({
      width: { size: scoreColW[i], type: WidthType.DXA },
      shading: { fill: NAVY, type: ShadingType.CLEAR },
      borders: { top: NIL, bottom: NIL, left: NIL, right: { style: BorderStyle.SINGLE, size: 4, color: WHITE } },
      margins: CM, verticalAlign: VerticalAlign.CENTER,
      children: [p([t(lbl, { bold: true, color: WHITE, size: 20 })], { alignment: AlignmentType.CENTER })]
    })),
    new TableCell({
      width: { size: OV_W, type: WidthType.DXA },
      shading: { fill: GOLD, type: ShadingType.CLEAR },
      borders: noB, margins: CM, verticalAlign: VerticalAlign.CENTER,
      children: [p([t('Overall Band', { bold: true, color: NAVY, size: 20 })], { alignment: AlignmentType.CENTER })]
    })
  ]});

  const scoreValueRow = new TableRow({ children: [
    ...skillScores.map((sc, i) => new TableCell({
      width: { size: scoreColW[i], type: WidthType.DXA },
      shading: { fill: LG, type: ShadingType.CLEAR },
      borders: { top: NIL, bottom: NIL, left: NIL, right: { style: BorderStyle.SINGLE, size: 4, color: BDR } },
      margins: CM, verticalAlign: VerticalAlign.CENTER,
      children: [p([t(String(sc || '—'), { bold: true, color: NAVY, size: 36 })], { alignment: AlignmentType.CENTER })]
    })),
    new TableCell({
      width: { size: OV_W, type: WidthType.DXA },
      shading: { fill: NAVY, type: ShadingType.CLEAR },
      borders: noB, margins: CM, verticalAlign: VerticalAlign.CENTER,
      children: [p([t(String(overall), { bold: true, color: GOLD, size: 48 })], { alignment: AlignmentType.CENTER })]
    })
  ]});

  const scoreTable = new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: scoreColW,
    rows: [scoreHeaderRow, scoreValueRow]
  });

  // 12. CEFR scale — 6 equal tiles
  const tileW = [1744, 1744, 1744, 1744, 1744, 1746];  // sum=10466
  const cefrRow = new TableRow({ children: CEFR_CODES.map((code, i) => {
    const active = (i === cefr.idx);
    return new TableCell({
      width: { size: tileW[i], type: WidthType.DXA },
      shading: { fill: active ? NAVY : LG, type: ShadingType.CLEAR },
      borders: { top: NIL, bottom: NIL, left: NIL, right: i < 5 ? { style: BorderStyle.SINGLE, size: 4, color: WHITE } : NIL },
      margins: { top: 120, bottom: 120, left: 100, right: 100 },
      verticalAlign: VerticalAlign.CENTER,
      children: [
        p([t(code, { bold: true, color: active ? GOLD : DARK, size: 20 })], { alignment: AlignmentType.CENTER }),
        p([t(CEFR_NAMES[i], { color: active ? WHITE : STXT, size: 17 })], { alignment: AlignmentType.CENTER })
      ]
    });
  })});
  const cefrScale = new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: tileW,
    rows: [cefrRow]
  });

  // 13. Signature + doc reference footer
  const footerRow = new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW * 0.6 | 0, CW - (CW * 0.6 | 0)],
    rows: [row([
      new TableCell({
        borders: { top: { style: BorderStyle.SINGLE, size: 4, color: BDR }, bottom: NIL, left: NIL, right: NIL },
        margins: { top: 120, bottom: 0, left: 0, right: 200 },
        children: [
          p([t('Academic Team Leader', { bold: true, size: 18, color: DARK })]),
          p([t('IELTS CITY', { size: 18, color: STXT })])
        ]
      }),
      new TableCell({
        borders: { top: { style: BorderStyle.SINGLE, size: 4, color: BDR }, bottom: NIL, left: NIL, right: NIL },
        margins: { top: 120, bottom: 0, left: 200, right: 0 },
        children: [p([t('Ref: ' + (admin.docRef || 'IC-OTR-' + (admin.testDate || '').replace(/\//g, '')), { size: 17, color: STXT })], { alignment: AlignmentType.RIGHT })]
      })
    ])]
  });

  return [
    headerBand, ep(6),
    disclaimer, ep(8),
    metaRow, ep(10),
    candTitle, ep(8),
    namePhotoRow, ep(8),
    divider,
    ep(8), dobRow, ep(8),
    countryTable, ep(10),
    resultTitle, ep(8),
    scoreTable, ep(8),
    cefrScale, ep(10),
    footerRow
  ];
}

// ── PAGE 2 — LISTENING SKILL BREAKDOWN ────────────────────────────────────
function buildPage2(data) {
  const { listening, scores } = data;
  const totalCorrect = (listening.answerMap || []).filter(Boolean).length;
  const byType = listening.byType || [];
  return [
    subHeader('LISTENING — SKILL BREAKDOWN'),
    ep(10),
    statCards(totalCorrect, scores.listening || '—'),
    ep(12),
    subHeader('Accuracy by Question Type'),
    ep(8),
    accuracyTable(byType),
    ep(12),
    subHeader('Full Answer Map'),
    ep(8),
    answerMap(listening.answerMap || [])
  ];
}

// ── PAGE 3 — READING SKILL BREAKDOWN ─────────────────────────────────────
function buildPage3(data) {
  const { reading, scores } = data;
  const totalCorrect = (reading.answerMap || []).filter(Boolean).length;
  const byType = reading.byType || [];
  return [
    subHeader('READING — SKILL BREAKDOWN'),
    ep(10),
    statCards(totalCorrect, scores.reading || '—'),
    ep(12),
    subHeader('Accuracy by Question Type'),
    ep(8),
    accuracyTable(byType),
    ep(12),
    subHeader('Full Answer Map'),
    ep(8),
    answerMap(reading.answerMap || [])
  ];
}

// ── PAGE 4 — SPEAKING FEEDBACK ────────────────────────────────────────────
function buildPage4(data) {
  const { speaking, scores } = data;
  const criteria = speaking.criteria || [];
  const items = [];
  items.push(skillBanner('SPEAKING', scores.speaking || '—'));
  items.push(ep(8));
  items.push(feedbackColHeader());
  criteria.forEach((cr, i) => {
    if (i > 0) items.push(ep(6));
    items.push(criterionTable(cr.name, cr.band, cr.subCriteria, S_LABELS));
  });
  return items;
}

// ── PAGE 5 — WRITING TASK 1 FEEDBACK ──────────────────────────────────────
function buildPage5(data) {
  const { writingTask1, scores } = data;
  const criteria = (writingTask1 && writingTask1.criteria) ? writingTask1.criteria : [];
  const band = (scores.writing || '—');
  const items = [];
  items.push(skillBanner('WRITING — Task 1', band));
  items.push(ep(8));
  items.push(feedbackColHeader());
  criteria.forEach((cr, i) => {
    if (i > 0) items.push(ep(6));
    items.push(criterionTable(cr.name, cr.band, cr.subCriteria, W1_LABELS));
  });
  return items;
}

// ── PAGE 6 — WRITING TASK 2 FEEDBACK ──────────────────────────────────────
function buildPage6(data) {
  const { writingTask2, scores } = data;
  const criteria = (writingTask2 && writingTask2.criteria) ? writingTask2.criteria : [];
  const band = (scores.writing || '—');
  const items = [];
  items.push(skillBanner('WRITING — Task 2', band));
  items.push(ep(8));
  items.push(feedbackColHeader());
  criteria.forEach((cr, i) => {
    if (i > 0) items.push(ep(6));
    items.push(criterionTable(cr.name, cr.band, cr.subCriteria, W2_LABELS));
  });
  return items;
}

// ── PAGE 7 — UNDERSTANDING YOUR RESULT ────────────────────────────────────
function buildPage7(data) {
  const { candidate, scores } = data;
  const overall = parseFloat(scores.overall || 0);
  const cefr    = cefrFromBand(overall);
  const desc    = getBandDescriptor(overall);

  // Title with gold bottom border
  const title = p([t('Understanding Your Result', { bold: true, color: NAVY, size: 40 })], {
    border: { bottom: { style: BorderStyle.SINGLE, size: 24, color: GOLD } },
    spacing: { before: 0, after: 200 }
  });

  // Result block: 3-col navy
  const rb1W = 1900, rb2W = 80, rb3W = CW - rb1W - rb2W;
  const resultBlock = new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [rb1W, rb2W, rb3W],
    rows: [row([
      // Col 1: OVERALL BAND + score
      navyCell([
        p([t('OVERALL BAND', { bold: true, color: GOLD, size: 18 })], { alignment: AlignmentType.CENTER }),
        ep(6),
        p([t(String(overall), { bold: true, color: WHITE, size: 72 })], { alignment: AlignmentType.CENTER })
      ], { width: { size: rb1W, type: WidthType.DXA }, margins: CML, verticalAlign: VerticalAlign.CENTER }),
      // Col 2: divider
      new TableCell({ width: { size: rb2W, type: WidthType.DXA }, shading: { fill: NAVY, type: ShadingType.CLEAR }, borders: noB, margins: CMZ, children: [ep()] }),
      // Col 3: YOUR RESULT + CEFR pill
      navyCell([
        p([t('YOUR RESULT', { bold: true, color: WHITE, size: 22 })]),
        ep(6),
        // CEFR pill — gold fill, 2500 DXA wide
        new Table({
          width: { size: 2500, type: WidthType.DXA },
          columnWidths: [2500],
          rows: [row([new TableCell({
            width: { size: 2500, type: WidthType.DXA },
            shading: { fill: GOLD, type: ShadingType.CLEAR },
            borders: noB,
            margins: { top: 80, bottom: 80, left: 160, right: 160 },
            children: [p([t(cefr.code + ' — ' + cefr.name, { bold: true, color: NAVY, size: 20 })], { alignment: AlignmentType.CENTER })]
          })])]
        })
      ], { width: { size: rb3W, type: WidthType.DXA }, margins: CML, verticalAlign: VerticalAlign.CENTER })
    ])]
  });

  // Skill chips — 4 equal cells (2587 + 40 + 2586 + 40 + 2587 + 40 + 2586 = 10466)
  const skillChipW = [2587, 40, 2586, 40, 2587, 40, 2586];
  const skillLabels  = ['L', 'R', 'W', 'S'];
  const skillScores  = [scores.listening, scores.reading, scores.writing, scores.speaking];
  const skillCols    = [0, 2, 4, 6];  // indices in chipW for content cells
  const gapCols      = [1, 3, 5];     // indices for gap cells

  const chipCells = skillChipW.map((w, i) => {
    const isGap = gapCols.includes(i);
    if (isGap) {
      return new TableCell({ width: { size: w, type: WidthType.DXA }, borders: noB, margins: CMZ, children: [ep()] });
    }
    const si = skillCols.indexOf(i);
    return new TableCell({
      width: { size: w, type: WidthType.DXA },
      shading: { fill: LG, type: ShadingType.CLEAR },
      borders: { top: { style: BorderStyle.SINGLE, size: 8, color: NAVY }, bottom: NIL, left: NIL, right: NIL },
      margins: { top: 120, bottom: 120, left: 160, right: 160 },
      verticalAlign: VerticalAlign.CENTER,
      children: [
        p([t(skillLabels[si], { bold: true, color: NAVY, size: 28 })], { alignment: AlignmentType.CENTER }),
        p([t(String(skillScores[si] || '—'), { color: STXT, size: 28 })], { alignment: AlignmentType.CENTER })
      ]
    });
  });
  const chipRow = new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: skillChipW,
    rows: [row(chipCells)]
  });

  // Descriptor block — full-width LB fill
  const descBlock = new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    rows: [row([new TableCell({
      width: { size: CW, type: WidthType.DXA },
      shading: { fill: LB, type: ShadingType.CLEAR },
      borders: noB,
      margins: { top: 160, bottom: 160, left: 240, right: 240 },
      children: [p([t(desc, { size: 18, color: DARK })])]
    })])]
  });

  // Footer note
  const footerNote = p([
    t('For queries regarding this report, please contact your IELTS CITY Academic Team Leader. This report remains the confidential property of IELTS CITY and is intended solely for the named candidate.', { size: 16, color: G88, italics: true })
  ]);

  return [
    title,
    ep(10),
    resultBlock,
    ep(10),
    chipRow,
    ep(10),
    descBlock,
    ep(10),
    footerNote
  ];
}

// ── DOCUMENT ASSEMBLER ────────────────────────────────────────────────────
async function buildDocument(data) {
  const hdr = makeHeader();
  const ftr = makeFooter();
  const pageProps = (m) => ({
    page: {
      size: { width: PW, height: PH, orientation: 'portrait' },
      margin: m
    }
  });

  const doc = new Document({
    creator: 'IELTS CITY OTR Tool',
    title: 'Official Test Report',
    sections: [
      // ── Section 1: Page 1 (M17 margins)
      {
        properties: pageProps(M17),
        headers: { default: hdr },
        footers: { default: ftr },
        children: buildPage1(data)
      },
      // ── Section 2: Pages 2–6 (M26 margins, PageBreak between each page)
      {
        properties: pageProps(M26),
        headers: { default: makeHeader() },
        footers: { default: makeFooter() },
        children: [
          ...buildPage2(data),
          pb(),
          ...buildPage3(data),
          pb(),
          ...buildPage4(data),
          pb(),
          ...buildPage5(data),
          pb(),
          ...buildPage6(data)
        ]
      },
      // ── Section 3: Page 7 (M17 margins)
      {
        properties: pageProps(M17),
        headers: { default: makeHeader() },
        footers: { default: makeFooter() },
        children: buildPage7(data)
      }
    ]
  });

  return Packer.toBuffer(doc);
}

// ── XML POST-PROCESSING ───────────────────────────────────────────────────
// Fix 1: pBdr element order — Word requires top/left/bottom/right in that order.
// The docx library may emit them out of order, causing Word to reject the file.
function fixPBdrOrder(xml) {
  return xml.replace(/<w:pBdr>[\s\S]*?<\/w:pBdr>/g, (match) => {
    const top   = match.match(/<w:top[^>]*\/>/);
    const left  = match.match(/<w:left[^>]*\/>/);
    const bot   = match.match(/<w:bottom[^>]*\/>/);
    const right = match.match(/<w:right[^>]*\/>/);
    if (!top || !left || !bot || !right) return match;
    const indM  = match.match(/\n( +)<w:/);
    const closeM = match.match(/( +)<\/w:pBdr>/);
    if (!indM || !closeM) return match;
    const ind   = indM[1];
    const close = closeM[1];
    return '<w:pBdr>\n' + ind + top[0] + '\n' + ind + left[0] + '\n' + ind + bot[0] + '\n' + ind + right[0] + '\n' + close + '</w:pBdr>';
  });
}

// Fix 2: Remove nested tblBorders inside table cells (causes rendering artefacts
// when nested tables are used for bars and chips).
function fixNestedTblBorders(xml) {
  // Match tblBorders elements that are deeply indented (≥ 12 spaces) — these are
  // nested inside cells, not the top-level table. Remove them entirely.
  return xml.replace(/\n {12,}<w:tblBorders>[\s\S]*?<\/w:tblBorders>/g, '');
}

// Fix 3: Inject photo placeholder floating shape.
// Inserts a WPS rectangle anchored near the candidate-name section.
// The shape is anchored relative to the column (H) and the paragraph (V).
const PHOTO_SHAPE = [
  '<w:p>',
    '<w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>',
    '<w:r>',
      '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">',
        '<mc:Choice Requires="wps">',
          '<w:drawing>',
            '<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
              ' distT="0" distB="0" distL="114300" distR="114300"',
              ' simplePos="0" relativeHeight="251659264" behindDoc="0"',
              ' locked="0" layoutInCell="1" allowOverlap="1">',
              '<wp:simplePos x="0" y="0"/>',
              '<wp:positionH relativeFrom="column"><wp:posOffset>5778240</wp:posOffset></wp:positionH>',
              '<wp:positionV relativeFrom="paragraph"><wp:posOffset>-8321</wp:posOffset></wp:positionV>',
              '<wp:extent cx="1143000" cy="1524000"/>',
              '<wp:effectExtent l="0" t="0" r="0" b="0"/>',
              '<wp:wrapNone/>',
              '<wp:docPr id="101" name="Photo Placeholder"/>',
              '<wp:cNvGraphicFramePr/>',
              '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
                '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">',
                  '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">',
                    '<wps:cNvPr id="101" name="Photo Placeholder"/>',
                    '<wps:cNvSpPr><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr>',
                    '<wps:spPr>',
                      '<a:xfrm><a:off x="0" y="0"/><a:ext cx="1143000" cy="1524000"/></a:xfrm>',
                      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
                      '<a:noFill/>',
                      '<a:ln w="3175"><a:solidFill><a:srgbClr val="AAAAAA"/></a:solidFill></a:ln>',
                    '</wps:spPr>',
                    '<wps:bodyPr>',
                      '<a:prstTxWarp prst="textBox"><a:avLst/></a:prstTxWarp>',
                    '</wps:bodyPr>',
                  '</wps:wsp>',
                '</a:graphicData>',
              '</a:graphic>',
            '</wp:anchor>',
          '</w:drawing>',
        '</mc:Choice>',
        '<mc:Fallback><w:pict/></mc:Fallback>',
      '</mc:AlternateContent>',
    '</w:r>',
  '</w:p>'
].join('');

function injectPhotoShape(xml) {
  // Anchor on ">Photo<" text node inside the placeholder cell.
  // From there, the 2nd </w:tbl> closes the namePhotoRow outer table.
  // docx generates compact XML (no newlines), so we match bare closing tags.
  const anchor = '>Photo<';
  const anchorIdx = xml.indexOf(anchor);
  if (anchorIdx === -1) return xml;

  const CLOSE_TBL = '</w:tbl>';
  let count    = 0;
  let injectAt = -1;
  for (let i = anchorIdx; i < xml.length - 8; i++) {
    if (xml.substring(i, i + 8) === CLOSE_TBL) {
      count++;
      if (count === 2) { injectAt = i + 8; break; }
    }
  }
  if (injectAt === -1) return xml;
  return xml.slice(0, injectAt) + PHOTO_SHAPE + xml.slice(injectAt);
}

async function applyXmlFixes(docxBuffer) {
  const zip = await JSZip.loadAsync(docxBuffer);
  let xml    = await zip.file('word/document.xml').async('text');

  xml = fixPBdrOrder(xml);
  xml = fixNestedTblBorders(xml);
  xml = injectPhotoShape(xml);

  zip.file('word/document.xml', xml);
  return zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 }
  });
}

// ── PAYLOAD ADAPTER ───────────────────────────────────────────────────────
// Normalise the Code.gs JSON payload into the shape expected by page builders.
function adaptPayload(raw) {
  // answerMap: object {q1:true, q2:false, ...} or boolean array → boolean array[40]
  function toAnswerArr(mapObj) {
    if (Array.isArray(mapObj)) return mapObj;
    const arr = Array(40).fill(false);
    if (!mapObj) return arr;
    Object.entries(mapObj).forEach(([k, v]) => {
      const idx = parseInt(k.replace(/\D/g, '')) - 1;
      if (idx >= 0 && idx < 40) arr[idx] = !!v;
    });
    return arr;
  }

  // byType: may come as array of {type, correct, total} or computed from answerMap
  function ensureByType(byType) {
    if (Array.isArray(byType) && byType.length > 0) return byType;
    return [];
  }

  const L = raw.listening  || {};
  const R = raw.reading    || {};
  const S = raw.speaking   || {};
  const W1 = raw.writingTask1 || {};
  const W2 = raw.writingTask2 || {};

  return {
    candidate: raw.candidate || {},
    admin:     raw.admin     || {},
    scores:    raw.scores    || {},
    listening: {
      answerMap: toAnswerArr(L.answerMap),
      byType:    ensureByType(L.byType)
    },
    reading: {
      answerMap: toAnswerArr(R.answerMap),
      byType:    ensureByType(R.byType)
    },
    speaking: {
      criteria: Array.isArray(S.criteria) ? S.criteria : []
    },
    writingTask1: {
      criteria: Array.isArray(W1.criteria) ? W1.criteria : []
    },
    writingTask2: {
      criteria: Array.isArray(W2.criteria) ? W2.criteria : []
    }
  };
}

// ── VERCEL HANDLER ────────────────────────────────────────────────────────
module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method Not Allowed' });
    return;
  }

  let data;
  try {
    data = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
    data = adaptPayload(data);
  } catch (e) {
    res.status(400).json({ error: 'Invalid JSON: ' + e.message });
    return;
  }

  try {
    const raw    = await buildDocument(data);
    const final  = await applyXmlFixes(raw);

    const name    = ((data.candidate.familyName || '') + '_' + (data.candidate.firstName || '') + '_OTR').trim().replace(/\s+/g, '_');
    const safeName = encodeURIComponent(name);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${name}.docx"; filename*=UTF-8''${safeName}.docx`);
    res.status(200).send(final);
  } catch (e) {
    console.error('[OTR] generation error:', e);
    res.status(500).json({ error: 'Generation failed: ' + e.message, stack: e.stack });
  }
};
