/**
 * ══════════════════════════════════════════════════════════════════════════════
 * GOOGLE APPS SCRIPT - BACKEND & CHẤM ĐIỂM BẢO MẬT (SERVER-SIDE GRADING)
 * Game: "Nhà Thám Hiểm Ngôn Từ – Giải Mã Mưa Mùa Xuân"
 *
 * 🔒 BẢO MẬT: Toàn bộ đáp án được lưu trữ và chấm điểm trực tiếp trên Server.
 * Học sinh KHÔNG THỂ xem trộm đáp án từ trình duyệt (F12 / View Source).
 * ══════════════════════════════════════════════════════════════════════════════
 */

// Tên các trang tính (sheet tabs)
const SHEET_SUBMISSIONS = 'Submissions';
const SHEET_ANSWERS = 'Answers';
const SHEET_CONFIG = 'Config';

// Mã PIN giáo viên mặc định
const DEFAULT_PIN = '123456';

// 5 Đáp án chuẩn mặc định
const DEFAULT_ANSWERS = [
  { q: 1, word: 'bé nhỏ', label: 'Từ miêu tả hạt mưa', hint: 'Hạt mưa xuân rất nhỏ mịn, mềm mại, rơi như nhảy nhót.' },
  { q: 2, word: 'khô cằn', label: 'Trạng thái mặt đất', hint: 'Mặt đất sau mùa đông lạnh khô nứt nẻ, khao khát được đánh thức.' },
  { q: 3, word: 'trong lành', label: 'Miêu tả giọt mưa', hint: 'Giọt mưa xuân trong trẻo, mang lại sự trong lành cho không gian.' },
  { q: 4, word: 'dịu mềm', label: 'Trạng thái sau mưa', hint: 'Mặt đất ngấm mưa xuân trở nên mềm mại, tơi xốp, màu mỡ.' },
  { q: 5, word: 'sức sống', label: 'Điều mưa mang cho cây', hint: 'Mưa xuân tiếp thêm nhựa sống và nguồn sức sống tràn trề cho cây cối.' }
];

/**
 * 1. HÀM KHỞI TẠO TỰ ĐỘNG BẢNG TÍNH (Chạy 1 lần duy nhất)
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Tạo sheet Submissions
  let subSheet = ss.getSheetByName(SHEET_SUBMISSIONS);
  if (!subSheet) {
    subSheet = ss.insertSheet(SHEET_SUBMISSIONS);
  }
  if (subSheet.getLastRow() === 0) {
    const headers = [
      'Thời gian',
      'Họ và tên',
      'Lớp',
      'Ô 1 (Hạt mưa)',
      'Ô 2 (Mặt đất)',
      'Ô 3 (Giọt mưa)',
      'Ô 4 (Sau mưa)',
      'Ô 5 (Mang lại)',
      'Số câu đúng',
      'Số câu sai',
      'Điểm số (/10)',
      'Tỉ lệ đúng (%)',
      'Lượt làm bài',
      'Thời gian làm bài'
    ];
    subSheet.appendRow(headers);
    const headerRange = subSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#368963');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    subSheet.setFrozenRows(1);
  } else if (subSheet.getLastColumn() < 14) {
    // Tự động bổ sung tiêu đề cột 14 nếu bảng đã tạo từ trước
    subSheet.getRange(1, 14).setValue('Thời gian làm bài')
      .setBackground('#368963')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }

  // 2. Tạo sheet Answers (Lưu đáp án trực tiếp trên Google Sheets)
  let ansSheet = ss.getSheetByName(SHEET_ANSWERS);
  if (!ansSheet) {
    ansSheet = ss.insertSheet(SHEET_ANSWERS);
    ansSheet.appendRow(['Câu', 'Vị trí câu', 'Đáp án đúng', 'Hướng dẫn / Giải thích']);
    const ansHeader = ansSheet.getRange(1, 1, 1, 4);
    ansHeader.setBackground('#087ba8');
    ansHeader.setFontColor('#ffffff');
    ansHeader.setFontWeight('bold');

    DEFAULT_ANSWERS.forEach(item => {
      ansSheet.appendRow([item.q, item.label, item.word, item.hint]);
    });
    ansSheet.setFrozenRows(1);
  }

  // 3. Tạo sheet Config (Mã PIN giáo viên)
  let configSheet = ss.getSheetByName(SHEET_CONFIG);
  if (!configSheet) {
    configSheet = ss.insertSheet(SHEET_CONFIG);
    configSheet.appendRow(['Tham số', 'Giá trị']);
    configSheet.appendRow(['teacher_pin', DEFAULT_PIN]);
    const cfgHeader = configSheet.getRange(1, 1, 1, 2);
    cfgHeader.setBackground('#cf6c32');
    cfgHeader.setFontColor('#ffffff');
    cfgHeader.setFontWeight('bold');
  }

  Logger.log('✅ Đã tạo cấu trúc bảng tính thành công (Submissions, Answers, Config)!');
}

/**
 * Lấy danh sách đáp án chuẩn từ sheet Answers (hoặc fallback mặc định)
 */
function getStandardAnswers() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ansSheet = ss.getSheetByName(SHEET_ANSWERS);
    if (!ansSheet || ansSheet.getLastRow() <= 1) {
      return DEFAULT_ANSWERS;
    }
    const data = ansSheet.getDataRange().getValues().slice(1);
    return data.map((row, idx) => ({
      q: Number(row[0]) || (idx + 1),
      label: String(row[1] || ''),
      word: String(row[2] || '').trim(),
      hint: String(row[3] || '')
    }));
  } catch (e) {
    return DEFAULT_ANSWERS;
  }
}

/**
 * Chuẩn hóa chuỗi so khớp tiếng Việt (Unicode NFC, lowercase, loại bỏ khoảng trắng thừa)
 */
function normalizeStr(s) {
  return (s || '')
    .toString()
    .trim()
    .normalize('NFC')
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

/**
 * Chuẩn hóa họ tên học sinh thành in hoa (Unicode NFC, loại bỏ khoảng trắng thừa)
 */
function normalizeName(s) {
  return (s || '')
    .toString()
    .trim()
    .normalize('NFC')
    .replace(/\s+/g, ' ')
    .toUpperCase();
}

/**
 * Định dạng số giây sang phút / giây (VD: 85 -> "1 phút 25 giây")
 */
function formatDuration(sec) {
  if (!sec || sec <= 0) return 'Dưới 1 giây';
  const m = Math.floor(sec / 60);
  const s = sec % 60;
  if (m > 0) {
    return `${m} phút ${s > 0 ? s + ' giây' : ''}`.trim();
  }
  return `${s} giây`;
}

/**
 * Chuyển chuỗi thời gian hoặc số giây thành số giây nguyên
 * Hỗ trợ: "1 phút 25 giây", "45 giây", "Dưới 1 giây", số nguyên,...
 */
function parseDurationToSeconds(val) {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : Math.round(val);
  const str = String(val).trim();
  if (str === '–' || str === '-' || !str) return 0;
  if (/^dưới\s*1\s*giây/i.test(str)) return 1;

  let totalSec = 0;
  const minMatch = str.match(/(\d+)\s*phút/i);
  const secMatch = str.match(/(\d+)\s*giây/i);

  if (minMatch) {
    totalSec += parseInt(minMatch[1], 10) * 60;
  }
  if (secMatch) {
    totalSec += parseInt(secMatch[1], 10);
  }
  if (!minMatch && !secMatch && !isNaN(Number(str))) {
    totalSec = parseInt(str, 10);
  }
  return totalSec;
}

/**
 * 2. NHẬN BÀI NỘP VÀ CHẤM ĐIỂM TRỰC TIẾP TRÊN SERVER (POST)
 */
function doPost(e) {
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    const postData = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : null;
    if (!postData) {
      return responseJSON({ success: false, error: 'Dữ liệu bài nộp trống.' });
    }

    const answersKey = getStandardAnswers();
    const studentAnswers = postData.answers || [
      postData.answer_1,
      postData.answer_2,
      postData.answer_3,
      postData.answer_4,
      postData.answer_5
    ];

    // CHẤM ĐIỂM BẢO MẬT TRÊN SERVER
    let correctCount = 0;
    const results = [];

    for (let i = 0; i < 5; i++) {
      const standard = answersKey[i] || DEFAULT_ANSWERS[i];
      const studentVal = studentAnswers[i] || '';
      const isRight = normalizeStr(studentVal) === normalizeStr(standard.word);

      if (isRight) {
        correctCount++;
        results.push({
          q: i + 1,
          isCorrect: true
        });
      } else {
        results.push({
          q: i + 1,
          isCorrect: false,
          userWord: studentVal,
          correctWord: standard.word,
          hint: standard.hint
        });
      }
    }

    const wrongCount = 5 - correctCount;
    const score = correctCount * 2;
    const accuracyPercent = correctCount * 20;

    // Lời động viên chuẩn theo mức điểm
    let encouragement = '';
    if (correctCount === 5) {
      encouragement = '🎉 Tuyệt vời! Bạn là nhà thám hiểm ngôn từ xuất sắc!';
    } else if (correctCount === 4) {
      encouragement = '👏 Giỏi lắm! Bạn gần hoàn hảo rồi!';
    } else if (correctCount === 3) {
      encouragement = '😊 Bạn đã tiến bộ rồi! Hãy xem lại những từ còn nhầm nhé.';
    } else if (correctCount >= 1) {
      encouragement = '💪 Cố gắng thêm một chút nhé! Bạn sẽ làm tốt hơn.';
    } else {
      encouragement = '🌱 Đừng nản nhé! Hãy đọc lại đoạn văn và thử lại lần nữa.';
    }

    // GHI DỮ LIỆU VÀO GOOGLE SHEETS
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let subSheet = ss.getSheetByName(SHEET_SUBMISSIONS);
    if (!subSheet) {
      setup();
      subSheet = ss.getSheetByName(SHEET_SUBMISSIONS);
    }

    // CHUẨN HOÁ TÊN IN HOA & KIỂM TRA HỌC SINH ĐÃ LÀM BÀI CHƯA
    const rawName = postData.student_name || 'Khuyết danh';
    const studentName = normalizeName(rawName);
    const studentClass = String(postData.student_class || '').trim();

    // Quét toàn bộ bảng Submissions để tìm số lượt đã làm trước đó của học sinh này
    let previousAttempts = 0;
    const lastRow = subSheet.getLastRow();
    if (lastRow > 1) {
      // Cột 2: Họ và tên, Cột 3: Lớp
      const studentRange = subSheet.getRange(2, 2, lastRow - 1, 2).getValues();
      for (let r = 0; r < studentRange.length; r++) {
        const rowName = normalizeName(studentRange[r][0]);
        const rowClass = String(studentRange[r][1] || '').trim();
        // So khớp theo Họ và tên in hoa và Lớp
        if (rowName === studentName && (!studentClass || !rowClass || rowClass === studentClass)) {
          previousAttempts++;
        }
      }
    }

    // Lượt làm bài thực tế: Tự động cộng thêm lượt nếu học sinh đã làm trước đó
    const currentAttempt = previousAttempts + 1;

    // Thời gian làm bài (từ lúc bấm bắt đầu đến khi nộp bài)
    const durationFormatted = postData.duration_formatted || (postData.duration_seconds ? formatDuration(Number(postData.duration_seconds)) : '–');

    const timestamp = postData.submitted_at || new Date().toISOString();
    const row = [
      timestamp,
      studentName, // Tên học sinh luôn được chuẩn hoá IN HOA
      studentClass,
      studentAnswers[0] || '',
      studentAnswers[1] || '',
      studentAnswers[2] || '',
      studentAnswers[3] || '',
      studentAnswers[4] || '',
      correctCount,
      wrongCount,
      score,
      accuracyPercent,
      currentAttempt, // Đã tự động cộng dồn lượt làm bài
      durationFormatted // Thời gian làm bài
    ];

    subSheet.appendRow(row);
    lock.releaseLock();

    // TRẢ KẾT QUẢ ĐÃ CHẤM VỀ CHO CLIENT (kèm lượt làm bài thực tế & thời gian làm bài)
    return responseJSON({
      success: true,
      score: score,
      correct_count: correctCount,
      wrong_count: wrongCount,
      accuracy_percent: accuracyPercent,
      encouragement: encouragement,
      attempt_number: currentAttempt,
      duration_formatted: durationFormatted,
      is_repeat_attempt: previousAttempts > 0,
      results: results
    });

  } catch (err) {
    return responseJSON({ success: false, error: err.toString() });
  }
}

/**
 * 3. PHỤC VỤ DASHBOARD GIÁO VIÊN (GET)
 */
function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const pin = params.pin || '';

    // Xác thực mã PIN giáo viên
    const validPin = getTeacherPin();
    if (pin !== validPin) {
      return responseJSON({ success: false, error: 'Mã PIN không chính xác.' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const subSheet = ss.getSheetByName(SHEET_SUBMISSIONS);

    if (!subSheet || subSheet.getLastRow() <= 1) {
      return responseJSON({
        success: true,
        submissions: [],
        stats: {
          total: 0,
          totalStudents: 0,
          totalSubmissions: 0,
          repeatCount: 0,
          avgScore: 0,
          avgAccuracy: 0,
          perfectCount: 0,
          avgDurationSec: 0,
          avgDuration: '–',
          questionStats: []
        }
      });
    }

    const data = subSheet.getDataRange().getValues();
    const rows = data.slice(1);

    const submissions = rows.map((r, idx) => ({
      stt: idx + 1,
      submitted_at: r[0],
      student_name: normalizeName(r[1]),
      student_class: String(r[2] || '').trim(),
      answer_1: r[3],
      answer_2: r[4],
      answer_3: r[5],
      answer_4: r[6],
      answer_5: r[7],
      correct_count: Number(r[8]) || 0,
      wrong_count: Number(r[9]) || 0,
      score: Number(r[10]) || 0,
      accuracy_percent: Number(r[11]) || 0,
      attempt_number: Number(r[12]) || 1,
      duration: r[13] || '–'
    }));

    // Nhóm theo học sinh duy nhất (Tên in hoa + Lớp) để KHÔNG tính lặp học sinh mới
    const studentMap = {};
    submissions.forEach(s => {
      const key = normalizeName(s.student_name) + '___' + s.student_class;
      if (!studentMap[key]) {
        studentMap[key] = {
          name: s.student_name,
          studentClass: s.student_class,
          attempts: [],
          bestScore: 0,
          latestScore: s.score
        };
      }
      studentMap[key].attempts.push(s);
      if (s.score > studentMap[key].bestScore) {
        studentMap[key].bestScore = s.score;
      }
      studentMap[key].latestScore = s.score;
    });

    const uniqueStudents = Object.values(studentMap);
    const totalStudents = uniqueStudents.length;
    const totalSubmissions = submissions.length;

    // Điểm trung bình và tỉ lệ tính theo điểm cao nhất của từng học sinh duy nhất
    const totalBestScore = uniqueStudents.reduce((sum, s) => sum + s.bestScore, 0);
    const avgScore = totalStudents > 0 ? (totalBestScore / totalStudents).toFixed(1) : 0;
    const avgAccuracy = totalStudents > 0 ? ((totalBestScore / (totalStudents * 10)) * 100).toFixed(0) : 0;
    const perfectCount = uniqueStudents.filter(s => s.bestScore === 10).length;
    const repeatCount = uniqueStudents.filter(s => s.attempts.length > 1).length;

    // Thống kê từng câu dựa trên đáp án chuẩn (tính trên tổng số lượt nộp)
    const answersKey = getStandardAnswers();
    const questionStats = answersKey.map((standard, qIdx) => {
      let qCorrect = 0;
      submissions.forEach(s => {
        const studentAns = normalizeStr(s['answer_' + (qIdx + 1)]);
        if (studentAns === normalizeStr(standard.word)) qCorrect++;
      });
      const qWrong = totalSubmissions - qCorrect;
      const correctRate = totalSubmissions > 0 ? Math.round((qCorrect / totalSubmissions) * 100) : 0;

      return {
        questionNum: qIdx + 1,
        label: standard.label,
        correctAnswer: standard.word,
        correctCount: qCorrect,
        wrongCount: qWrong,
        correctRate: correctRate,
        wrongRate: 100 - correctRate
      };
    });

    // Thời gian làm bài trung bình (tính trên các lượt nộp có ghi nhận thời gian > 0)
    const validDurations = submissions
      .map(s => parseDurationToSeconds(s.duration))
      .filter(sec => sec > 0);
    const avgDurationSec = validDurations.length > 0
      ? Math.round(validDurations.reduce((sum, d) => sum + d, 0) / validDurations.length)
      : 0;
    const avgDuration = avgDurationSec > 0 ? formatDuration(avgDurationSec) : '–';

    return responseJSON({
      success: true,
      submissions: submissions,
      stats: {
        total: totalStudents, // Số học sinh duy nhất (không bị lặp)
        totalStudents: totalStudents,
        totalSubmissions: totalSubmissions,
        repeatCount: repeatCount,
        avgScore: avgScore,
        avgAccuracy: avgAccuracy,
        perfectCount: perfectCount,
        avgDurationSec: avgDurationSec,
        avgDuration: avgDuration,
        questionStats: questionStats
      }
    });

  } catch (err) {
    return responseJSON({ success: false, error: err.toString() });
  }
}

/**
 * Lấy mã PIN từ sheet Config
 */
function getTeacherPin() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cfg = ss.getSheetByName(SHEET_CONFIG);
    if (!cfg) return DEFAULT_PIN;
    const data = cfg.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === 'teacher_pin') {
        return String(data[i][1]).trim();
      }
    }
    return DEFAULT_PIN;
  } catch (e) {
    return DEFAULT_PIN;
  }
}

/**
 * Trả về JSON chuẩn với ContentService
 */
function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
