// === НАСТРОЙКИ ===
const OPENAI_API_KEY = 'YOUR_OPENAI_API_KEY'; // Замени на свой ключ
const MODEL = 'gpt-4o-mini';

// === ГЛАВНАЯ ФУНКЦИЯ ===
function analyzeReviews() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  // Проходим по всем отзывам (начиная со 2-й строки)
  for (let i = 2; i <= lastRow; i++) {
    const review = sheet.getRange(i, 1).getValue();
    
    // Пропускаем пустые ячейки или уже обработанные
    if (!review || sheet.getRange(i, 2).getValue()) continue;
    
    // Анализируем отзыв
    const analysis = analyzeWithGPT(review);
    
    if (analysis) {
      sheet.getRange(i, 2).setValue(analysis.tonality);      // Тональность
      sheet.getRange(i, 3).setValue(analysis.category);      // Категория
      sheet.getRange(i, 4).setValue(analysis.recommendation); // Рекомендация
      sheet.getRange(i, 5).setValue(new Date());             // Дата анализа
    }
    
    // Пауза между запросами (чтобы не превысить лимит API)
    Utilities.sleep(500);
  }
  
  // Применяем цветовое форматирование
  applyConditionalFormatting();
  
  SpreadsheetApp.getUi().alert('Анализ завершён!');
}

// === ВЫЗОВ OPENAI API ===
function analyzeWithGPT(review) {
  const prompt = `Проанализируй отзыв клиента и верни JSON:

Отзыв: "${review}"

Верни ТОЛЬКО JSON без пояснений:
{
  "tonality": "positive" или "neutral" или "negative",
  "category": одна из категорий: "качество", "сервис", "цена", "скорость", "техподдержка",
  "recommendation": краткая рекомендация по улучшению (1 предложение)
}`;

  const payload = {
    model: MODEL,
    messages: [
      { role: 'system', content: 'Ты аналитик клиентских отзывов. Отвечай только JSON.' },
      { role: 'user', content: prompt }
    ],
    temperature: 0.3
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + OPENAI_API_KEY },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
    const json = JSON.parse(response.getContentText());
    const content = json.choices[0].message.content;
    
    // Парсим JSON из ответа
    return JSON.parse(content);
  } catch (e) {
    Logger.log('Ошибка: ' + e.message);
    return null;
  }
}

// === ЦВЕТОВОЕ ФОРМАТИРОВАНИЕ ===
function applyConditionalFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(2, 2, lastRow - 1, 1); // Столбец B (Тональность)
  
  // Очищаем старые правила
  sheet.clearConditionalFormatRules();
  
  const rules = [];
  
  // Зелёный для positive
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('positive')
    .setBackground('#d4edda')
    .setRanges([range])
    .build());
  
  // Жёлтый для neutral
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('neutral')
    .setBackground('#fff3cd')
    .setRanges([range])
    .build());
  
  // Красный для negative
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('negative')
    .setBackground('#f8d7da')
    .setRanges([range])
    .build());
  
  sheet.setConditionalFormatRules(rules);
}

// === МЕНЮ В ТАБЛИЦЕ ===
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔍 Анализ отзывов')
    .addItem('Запустить анализ', 'analyzeReviews')
    .addToUi();
}
