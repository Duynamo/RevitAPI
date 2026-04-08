/*
 * Google Apps Script - AI Translator (VN -> JP)
 * BẢN CẬP NHẬT MỚI: Dùng GROQ API (Model Llama siêu tốc) để tránh giới hạn
 */

// Chạy hàm này khi mở Google Sheet để tạo menu
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 AI Translator')
      .addItem('1. 🔑 Thiết lập Groq API Key', 'setApiKey')
      .addItem('2. 🇯🇵 Dịch Toàn Bộ Sheet (VN -> JP)', 'translateCurrentSheet')
      .addToUi();
}

// Lưu API Key vào Properties Script (Bảo mật, mỗi user tự set)
function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Cài đặt Groq API Key',
    'Nhập Groq API Key của bạn (Lấy miễn phí từ console.groq.com):\n(Nếu đã nhập rồi thì việc nhập lại sẽ tự ghi đè API cũ)',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.OK) {
    const key = result.getResponseText().trim();
    if (key !== '') {
      PropertiesService.getScriptProperties().setProperty('GROQ_API_KEY', key);
      ui.alert('Thành công', 'Đã lưu API Key của Groq an toàn.', ui.ButtonSet.OK);
    }
  }
}

// Lấy API key đã lưu
function getApiKey() {
  return PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY');
}

// Hàm chính để dịch toàn bộ Sheet
function translateCurrentSheet() {
  const apiKey = getApiKey();
  const ui = SpreadsheetApp.getUi();

  if (!apiKey) {
    ui.alert('Chưa có API Key', 'Vui lòng nhấn vào "Thiết lập Groq API Key" để khai báo trước khi sử dụng.', ui.ButtonSet.OK);
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getDataRange();
  const values = range.getValues();
  
  if (values.length === 0 || (values.length === 1 && values[0].length === 1 && values[0][0] === '')) {
    ui.alert('Lỗi', 'Sheet hiện tại đang trống.', ui.ButtonSet.OK);
    return;
  }

  // Xác nhận từ người dùng
  const response = ui.alert(
    'Xác nhận dịch',
    'Chương trình hiện đang sử dụng AI từ Groq (Siêu Tốc). Dữ liệu sẽ tự động gom nhóm để dịch hàng loạt.\n\nDữ liệu sau khi dịch xong sẽ được dán vào một Sheet mới. Tiếp tục?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Đang thu thập dữ liệu...", "AI Translator", -1);

  // 1. Thu thập các ô cần dịch để gom thành Nhóm (Batch)
  let cellsToTranslate = []; // lưu trữ { r, c, text, translated, error }
  let newValues = JSON.parse(JSON.stringify(values)); // deep copy dữ liệu

  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      let val = values[i][j];
      if (typeof val === 'string' && val.trim() !== '') {
        cellsToTranslate.push({ r: i, c: j, text: val.trim(), translated: null, error: null });
      }
    }
  }

  if (cellsToTranslate.length === 0) {
    ui.alert('Thông báo', 'Toàn bộ sheet đều là số hoặc trống, không có chữ nào để dịch.', ui.ButtonSet.OK);
    return;
  }

  // 2. Chia Batch và gọi AI 
  const BATCH_SIZE = 15; // Giữ nguyên mức 15 để an toàn với token limit
  let totalBatches = Math.ceil(cellsToTranslate.length / BATCH_SIZE);

  for (let b = 0; b < cellsToTranslate.length; b += BATCH_SIZE) {
    let batch = cellsToTranslate.slice(b, b + BATCH_SIZE);
    let originalTexts = batch.map(item => item.text);
    let currentBatchNum = Math.floor(b / BATCH_SIZE) + 1;
    
    ss.toast(`Hệ thống đang dịch yêu cầu cụm ${currentBatchNum} / ${totalBatches}...`, "AI Translator", -1);
    
    let success = false;
    let retryWaitTime = 6000; // API Groq Free Rate Limit hơi khắt khe nếu bắn liên tục, lùi sâu hơn 1 chút
    let translatedArray = [];
    
    // Thử tối đa 3 lần cho mỗi lô 
    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        translatedArray = callGroqBatchAPI(originalTexts, apiKey);
        
        // Kiểm tra an toàn
        if (!Array.isArray(translatedArray) || translatedArray.length !== originalTexts.length) {
          throw new Error(`AI trả về số lượng lệch (${translatedArray.length} kết quả so với ${originalTexts.length} gốc)`);
        }
        
        success = true;
        break; // Dịch thành công thì bẻ loop retry để sang lô tiếp theo
      } catch (err) {
        Logger.log(`Batch ${currentBatchNum} bị lỗi (lần ${attempt}/3): ${err.message}`);
        
        // Block app nếu API Key bị sai hoặc chưa có
        if (err.message.includes("API_KEY") || err.message.toLowerCase().includes("api key") || err.message.includes("unauthorized") || err.message.includes("key error")) {
           ui.alert('Lỗi Groq API Key', 'Mã xác thực của bạn không hợp lệ. Vui lòng lấy lại Keyword trên console.groq.com. \nChi tiết lỗi: ' + err.message, ui.ButtonSet.OK);
           return;
        }

        if (attempt < 3) {
          ss.toast(`Cụm ${currentBatchNum} bị quá tải Groq API. Nghỉ ${retryWaitTime/1000}s rồi vọt lại lần ${attempt+1}...`, "AI Translator", -1);
          Utilities.sleep(retryWaitTime);
          retryWaitTime += 4000; 
        } else {
          // Lỗi liên tục 3 lần -> Bỏ qua và báo thẳng vào ô
          for (let k = 0; k < batch.length; k++) {
             batch[k].error = `[LỖI QUÁ TẢI API SAU 3 LẦN THỬ: ${err.message}]`;
          }
        }
      }
    }

    if (success) {
      for (let k = 0; k < batch.length; k++) {
        batch[k].translated = translatedArray[k];
      }
    }

    Utilities.sleep(3000); // Khoảng dừng giữa các cụm (rate limit bảo toàn cho free tier của Groq)
  }

  // 3. Đổ dữ liệu đã xử lý vào Array 2D kết quả
  for (let item of cellsToTranslate) {
    if (item.error) {
      newValues[item.r][item.c] = item.error + "\n\nBản gốc: " + item.text;
    } else if (item.translated) {
      newValues[item.r][item.c] = item.translated;
    }
  }

  // 4. Sinh Sheet Mới để lưu data
  const newSheetName = sheet.getName() + "_JP";
  let newSheet = ss.getSheetByName(newSheetName);
  
  if (!newSheet) {
    newSheet = ss.insertSheet(newSheetName);
  } else {
    newSheet.clear(); 
  }

  newSheet.getRange(1, 1, newValues.length, newValues[0].length).setValues(newValues);
  for (let col = 1; col <= newValues[0].length; col++) {
    newSheet.setColumnWidth(col, sheet.getColumnWidth(col));
  }

  ss.toast("Thuật toán đã hoàn thành. Hãy kiểm tra kết quả ngay!", "✅ AI Translator", -1);
}

// -----------------------------------------------------
// Hàm Lõi Gọi Phương Thức Của Groq 
// -----------------------------------------------------
function callGroqBatchAPI(textsArray, apiKey) {
  const url = `https://api.groq.com/openai/v1/chat/completions`;
  
  // Mẹo: Đưa array thành object có ID để ép AI trả về đúng số lượng và không gộp/tách câu
  let inputObj = {};
  for (let i = 0; i < textsArray.length; i++) {
    inputObj[`id_${i}`] = textsArray[i];
  }
  
  const prompt = `Dịch các đoạn văn tiếng Việt sau sang tiếng Nhật.
Yêu cầu BẮT BUỘC:
1. Trả về đúng MỘT ĐỐI TƯỢNG JSON.
2. Đối tượng JSON kết quả PHẢI có CHÍNH XÁC các key tương tự như đầu vào (id_0, id_1, v.v...). 
3. Giá trị của mỗi key là văn bản tiếng Nhật tương ứng. Tuyệt đối không tự gộp hay tách các key.
4. Lời dịch tự nhiên, lịch sự (thể Masu/Desu).
5. Trả về chuẩn JSON, không có mã Markdown.

Dữ liệu đầu vào:
${JSON.stringify(inputObj, null, 2)}`;

  const payload = {
    // Model Llama 3.3 đang là đỉnh phong của Groq hiện tại
    "model": "llama-3.3-70b-versatile", 
    "messages": [
      {
        "role": "system",
        "content": "You are a professional Japanese translator system. You strictly reply with valid JSON objects matching the structural keys of the input."
      },
      {
        "role": "user",
        "content": prompt
      }
    ],
    "temperature": 0.1,
    // Bắt buộc ép model của GROQ phải móc ra cấu trúc JSON
    "response_format": { "type": "json_object" } 
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "headers": {
      "Authorization": "Bearer " + apiKey
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const contentText = response.getContentText();
  let json;
  
  try {
    json = JSON.parse(contentText);
  } catch(e) {
    throw new Error("Dữ liệu xuất ra từ Groq không thuộc chuẩn JSON: " + contentText.substring(0, 50));
  }

  if (responseCode === 200) {
    if (json.choices && json.choices.length > 0) {
       let resultRawText = json.choices[0].message.content;
       try {
           let parsedObject = JSON.parse(resultRawText); 
           
           // Khôi phục mảng từ Object
           let resultArray = [];
           for (let i = 0; i < textsArray.length; i++) {
             let key = `id_${i}`;
             if (parsedObject[key] !== undefined && parsedObject[key] !== null) {
               resultArray.push(String(parsedObject[key]));
             } else {
               throw new Error(`AI đánh rơi mất đoạn dịch của khoá [${key}].`);
             }
           }
           
           return resultArray;
       } catch (err) {
           throw new Error("Kết quả từ Groq bị sai cấu trúc JSON id hoặc thiếu key. Lỗi: " + err.message);
       }
    } else {
       throw new Error("Không tìm thấy nội dung phản hồi từ mảng choice của Groq.");
    }
  } else {
    throw new Error(json.error ? json.error.message : "Mã lỗi từ máy chủ Groq " + responseCode);
  }
}
