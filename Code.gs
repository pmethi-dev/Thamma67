const SPREADSHEET_ID = '1oU2-afmiE_hiO2n2aqsiP0nH2NlO_OK1ubfDsum3omQ';
const SALT = 'DhammaQuiz_Salt_2026_SecureKey';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('คลังข้อสอบ ธรรมศึกษา')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function hashPassword(password) {
  if (!password) return '';
  const saltedPassword = password.toString() + SALT;
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, saltedPassword, Utilities.Charset.UTF_8);
  return rawHash.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function getUsersSheet() {
  const ss = getSS();
  let sheet = ss.getSheetByName('users');
  if (!sheet) {
    sheet = ss.insertSheet('users');
    sheet.appendRow(['Username', 'Password', 'FullName', 'Nickname', 'Classroom', 'No', 'School']);
  }
  return sheet;
}

function getScoresSheet() {
  const ss = getSS();
  let sheet = ss.getSheetByName('scores');
  if (!sheet) {
    sheet = ss.insertSheet('scores');
    sheet.appendRow(['Timestamp', 'FullName', 'Nickname', 'Classroom', 'No', 'School', 'Level', 'Subject', 'Score', 'Total', 'Percentage', 'Username']);
  }
  return sheet;
}

function cleanClassroomString(val) {
  if (!val) return '-';
  const str = val.toString().trim();
  if (str.includes('GMT') || str.includes('00:00:00')) {
    const matched = str.match(/^(.*?)(?=\s[A-Z][a-z]{2}\s[A-Z][a-z]{2})/);
    return matched && matched[1] ? matched[1].trim() : 'ไม่ระบุ';
  }
  return str;
}

function getAllQuizData() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('all_quiz_data_v1');
  if (cached) return JSON.parse(cached);

  try {
    const ss = getSS();
    const result = {};
    const sheets = ['tri', 'tho', 'ek'];

    sheets.forEach(level => {
      const sheet = ss.getSheetByName(level);
      if (!sheet) {
        result[level] = [];
        return;
      }
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) {
        result[level] = [];
        return;
      }

      const list = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row[1] || row[1].toString().trim() === '') continue;

        let rawCategory = row[0] ? row[0].toString().trim() : '';
        let subject = 'ธรรมวิภาค';
        if (rawCategory.includes('พุทธ') || rawCategory.includes('ประวัติ')) {
          subject = 'พุทธประวัติ';
        } else if (rawCategory.includes('วินัย') || rawCategory.includes('เบญจศีล') || rawCategory.includes('เบญจธรรม')) {
          subject = 'วินัย';
        }

        list.push({
          id: i,
          subject: subject,
          category: rawCategory || 'ทั่วไป',
          title: row[1].toString(),
          options: [
            row[2] ? row[2].toString() : '',
            row[3] ? row[3].toString() : '',
            row[4] ? row[4].toString() : '',
            row[5] ? row[5].toString() : ''
          ]
        });
      }
      result[level] = list;
    });

    cache.put('all_quiz_data_v1', JSON.stringify(result), 3600);
    return result;
  } catch (e) {
    return { tri: [], tho: [], ek: [] };
  }
}

function submitAndGradeQuiz(payload) {
  try {
    const { levelKey, userAnswers, username, subject } = payload;
    const userSheet = getUsersSheet();
    const userRows = userSheet.getDataRange().getValues();
    let currentUserData = null;

    for (let i = 1; i < userRows.length; i++) {
      if (userRows[i][0].toString().toLowerCase().trim() === username.toLowerCase().trim()) {
        currentUserData = {
          username: userRows[i][0],
          fullname: userRows[i][2],
          nickname: userRows[i][3],
          classroom: cleanClassroomString(userRows[i][4]),
          no: userRows[i][5],
          school: userRows[i][6]
        };
        break;
      }
    }

    if (!currentUserData) return { status: 'error', message: 'ไม่พบข้อมูลผู้ใช้งาน' };

    const ss = getSS();
    const sheet = ss.getSheetByName(levelKey);
    if (!sheet) return { status: 'error', message: 'ไม่พบระดับชั้นที่เลือก' };

    const data = sheet.getDataRange().getValues();
    let score = 0;

    const gradedResults = userAnswers.map(item => {
      const rowIndex = item.id;
      const row = data[rowIndex];
      if (!row) return { id: item.id, isCorrect: false, correctAnswerIndex: 0, correctAnswerText: '', exp: '' };

      const realAnswer = (row[6] !== "" && row[6] !== null) ? parseInt(row[6], 10) : 0;
      const isCorrect = (item.selectedIndex === realAnswer);
      if (isCorrect) score++;

      return {
        id: item.id,
        isCorrect: isCorrect,
        correctAnswerIndex: realAnswer,
        correctAnswerText: row[2 + realAnswer] ? row[2 + realAnswer].toString() : '',
        exp: row[7] ? row[7].toString() : ''
      };
    });

    const total = userAnswers.length;

    saveUserScoreInternal({
      username: currentUserData.username,
      fullname: currentUserData.fullname,
      nickname: currentUserData.nickname,
      classroom: currentUserData.classroom,
      no: currentUserData.no,
      school: currentUserData.school,
      level: levelKey,
      subject: subject,
      score: score,
      total: total
    });

    return { status: 'success', score: score, total: total, details: gradedResults };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function saveUserScoreInternal(userData) {
  try {
    const sheet = getScoresSheet();
    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy, HH:mm:ss");
    const scoreVal = parseInt(userData.score, 10) || 0;
    const totalVal = parseInt(userData.total, 10) || 0;
    const percentage = totalVal > 0 ? parseFloat(((scoreVal / totalVal) * 100).toFixed(1)) : 0;

    let cleanLevel = userData.level || '-';
    const levelStr = cleanLevel.toString().toUpperCase();
    if (levelStr.includes('TRI') || levelStr.includes('ตรี')) cleanLevel = 'ชั้นตรี';
    else if (levelStr.includes('THO') || levelStr.includes('โท')) cleanLevel = 'ชั้นโท';
    else if (levelStr.includes('EK') || levelStr.includes('เอก')) cleanLevel = 'ชั้นเอก';

    sheet.appendRow([
      timestamp,
      userData.fullname || '-',
      userData.nickname || '-',
      userData.classroom || '-',
      userData.no || '-',
      userData.school || '-',
      cleanLevel,
      userData.subject || '-',
      scoreVal,
      totalVal,
      percentage,
      userData.username || '-'
    ]);

    CacheService.getScriptCache().remove('leaderboard_data_v2');
  } catch (e) {
    Logger.log("Error saveUserScoreInternal: " + e.toString());
  }
}

function getUserDashboardData(username) {
  try {
    if (!username || username.toString().trim() === '') {
      return { status: 'error', message: 'กรุณาเข้าสู่ระบบใหม่อีกครั้ง' };
    }

    const cleanUser = username.toString().toLowerCase().trim();
    const userSheet = getUsersSheet();
    const userRows = userSheet.getDataRange().getValues();
    let profile = null;

    for (let i = 1; i < userRows.length; i++) {
      const uName = userRows[i][0] ? userRows[i][0].toString().toLowerCase().trim() : '';
      if (uName === cleanUser) {
        profile = {
          username: userRows[i][0].toString().trim(),
          fullname: userRows[i][2] ? userRows[i][2].toString().trim() : '',
          nickname: userRows[i][3] ? userRows[i][3].toString().trim() : '',
          classroom: cleanClassroomString(userRows[i][4]),
          no: userRows[i][5] ? userRows[i][5].toString().trim() : '',
          school: userRows[i][6] ? userRows[i][6].toString().trim() : ''
        };
        break;
      }
    }

    if (!profile) return { status: 'error', message: 'ไม่พบข้อมูลผู้ใช้งาน' };

    const scoreSheet = getScoresSheet();
    const scoreRows = scoreSheet.getDataRange().getValues();

    let history = [];
    let accumulatedScore = 0;
    let qualifyingCount = 0; // 🔥 จำนวนครั้งที่สอบผ่านเกณฑ์ (เอาไปนับเป็นไฟ)

    for (let i = 1; i < scoreRows.length; i++) {
      const row = scoreRows[i];
      if (!row || !row[0]) continue;

      const rowUsername = row[11] ? row[11].toString().toLowerCase().trim() : '';
      const rowFullname = row[1] ? row[1].toString().toLowerCase().trim() : '';

      const isMatch = (rowUsername !== '' && rowUsername === cleanUser) ||
        (rowFullname !== '' && profile.fullname !== '' && rowFullname === profile.fullname.toLowerCase());

      if (isMatch) {
        const score = parseInt(row[8], 10) || 0;
        const total = parseInt(row[9], 10) || 0;
        let percent = row[10] ? parseFloat(row[10]) : (total > 0 ? (score / total) * 100 : 0);
        if (isNaN(percent)) percent = 0;

        const timeStr = row[0] instanceof Date
          ? Utilities.formatDate(row[0], "GMT+7", "dd/MM/yyyy, HH:mm:ss")
          : row[0].toString();

        history.push({
          timestamp: timeStr,
          level: row[6] ? row[6].toString() : '-',
          subject: row[7] ? row[7].toString() : '-',
          score: score,
          total: total,
          percent: parseFloat(percent.toFixed(1))
        });

        accumulatedScore += score;

        // 🛡️ เกณฑ์การได้ไฟ: ต้องได้คะแนนอย่างน้อย 30% ขึ้นไปเท่านั้น ถึงจะนับเพิ่มไฟ +1
        if (percent >= 30) {
          qualifyingCount++;
        }
      }
    }

    history.reverse();

    return {
      status: 'success',
      profile: profile,
      stats: {
        totalAttempts: history.length,         // จำนวนครั้งที่ทำทั้งหมด
        qualifyingCount: qualifyingCount,     // 🔥 จำนวนไฟ (ทำผ่านเกณฑ์ 50% ขึ้นไป)
        accumulatedScore: accumulatedScore
      },
      history: history
    };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function registerUser(data) {
  try {
    const sheet = getUsersSheet();
    const rows = sheet.getDataRange().getValues();
    const username = data.username.toLowerCase().trim();

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0].toString().toLowerCase().trim() === username) {
        return { status: 'error', message: 'Username นี้ถูกใช้ไปแล้ว' };
      }
    }

    const hashedPassword = hashPassword(data.password);
    sheet.appendRow([username, hashedPassword, data.fullname, data.nickname, data.classroom, data.no, data.school]);

    return {
      status: 'success',
      user: { username: username, fullname: data.fullname, nickname: data.nickname, classroom: data.classroom, no: data.no, school: data.school }
    };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function loginUser(username, password) {
  try {
    const sheet = getUsersSheet();
    const rows = sheet.getDataRange().getValues();
    const targetUser = username.toLowerCase().trim();
    const inputPassword = password.toString();
    const inputHashedPassword = hashPassword(inputPassword);

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0].toString().toLowerCase().trim() === targetUser) {
        const storedPassword = rows[i][1].toString();
        const isHashMatch = storedPassword === inputHashedPassword;
        const isPlainMatch = storedPassword === inputPassword;

        if (isHashMatch || isPlainMatch) {
          if (isPlainMatch && !isHashMatch) {
            sheet.getRange(i + 1, 2).setValue(inputHashedPassword);
          }
          return {
            status: 'success',
            user: { username: rows[i][0], fullname: rows[i][2], nickname: rows[i][3], classroom: cleanClassroomString(rows[i][4]), no: rows[i][5], school: rows[i][6] }
          };
        } else {
          return { status: 'error', message: 'รหัสผ่านไม่ถูกต้อง' };
        }
      }
    }
    return { status: 'error', message: 'ไม่พบชื่อผู้ใช้งานนี้' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function updateUserProfile(data) {
  try {
    const sheet = getUsersSheet();
    const rows = sheet.getDataRange().getValues();
    const targetUser = data.username.toLowerCase().trim();

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0].toString().toLowerCase().trim() === targetUser) {
        sheet.getRange(i + 1, 3).setValue(data.fullname);
        sheet.getRange(i + 1, 4).setValue(data.nickname);
        sheet.getRange(i + 1, 5).setValue(data.classroom);
        sheet.getRange(i + 1, 6).setValue(data.no);
        sheet.getRange(i + 1, 7).setValue(data.school);

        if (data.password && data.password.toString().trim() !== '') {
          sheet.getRange(i + 1, 2).setValue(hashPassword(data.password));
        }

        return {
          status: 'success',
          user: { username: rows[i][0], fullname: data.fullname, nickname: data.nickname, classroom: data.classroom, no: data.no, school: data.school }
        };
      }
    }
    return { status: 'error', message: 'ไม่พบผู้ใช้' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function getLeaderboard() {
  try {
    const cache = CacheService.getScriptCache();
    const cachedData = cache.get('leaderboard_data_v2');
    if (cachedData !== null) return JSON.parse(cachedData);

    const ss = getSS();
    const sheet = ss.getSheetByName('scores');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    data.shift();

    const rawList = data.map(row => {
      const score = parseInt(row[8], 10) || 0;
      const total = parseInt(row[9], 10) || 0;
      let percent = (row[10] !== "" && row[10] !== undefined) ? parseFloat(row[10]) : (total > 0 ? (score / total) * 100 : 0);

      let rawLevel = row[6] ? row[6].toString().trim() : '-';
      let cleanLevel = rawLevel;
      if (rawLevel.includes('TRI') || rawLevel.includes('ตรี')) cleanLevel = 'ชั้นตรี';
      else if (rawLevel.includes('THO') || rawLevel.includes('โท')) cleanLevel = 'ชั้นโท';
      else if (rawLevel.includes('EK') || rawLevel.includes('เอก')) cleanLevel = 'ชั้นเอก';

      return {
        fullname: row[1] ? row[1].toString().trim() : '-',
        nickname: row[2] ? row[2].toString().trim() : '-',
        classroom: cleanClassroomString(row[3]),
        no: row[4] ? row[4].toString() : '-',
        school: row[5] ? row[5].toString() : '-',
        level: cleanLevel,
        subject: row[7] ? row[7].toString() : '-',
        score: score,
        total: total,
        percent: parseFloat(percent.toFixed(1))
      };
    });

    rawList.sort((a, b) => b.score - a.score || b.percent - a.percent);

    const uniqueMap = new Map();
    rawList.forEach(item => {
      const uniqueKey = `${item.fullname}_${item.school}_${item.level}`.toLowerCase();
      if (!uniqueMap.has(uniqueKey)) {
        uniqueMap.set(uniqueKey, item);
      }
    });

    const uniqueList = Array.from(uniqueMap.values());
    cache.put('leaderboard_data_v2', JSON.stringify(uniqueList), 300);
    return uniqueList;
  } catch (e) {
    return [];
  }
}
