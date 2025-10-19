const SHEET_NAME = 'EED02 Clips';
const OUTPUT_STATUS_CELL = 'E9:F12';   // Merged cell
const OUTPUT_TEXT_CELL = 'E14:F29';    // Merged cell
const HASH_STORE_CELL = 'G8';
const API_URL = 'https://clip-combo-flyio.fly.dev/solve';

function runCombinationSolver() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const now = new Date();
  const startTimestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMM d, yyyy - h:mm:ss a');
  const estRuntime = '~90s (may vary)';
  const runningMsg = `\n⏳ Running...\nStarted at ${startTimestamp}`;

  appendToCell(sheet, OUTPUT_STATUS_CELL, runningMsg);

  try {
    const halfClipData = sheet.getRange('A3:B').getValues().filter(r => r[0] !== '');
    const wholeClipData = sheet.getRange('C3:D').getValues().filter(r => r[0] !== '');
    const z = parseInt(sheet.getRange('F2').getValue());
    const num_half = parseInt(sheet.getRange('F3').getValue());
    const num_whole = parseInt(sheet.getRange('F4').getValue());
    const tolerance = parseFloat(sheet.getRange('F5').getValue());

    const half_clips = flattenInventory(halfClipData);
    const whole_clips = flattenInventory(wholeClipData);

    const input = {
      whole_clips,
      half_clips,
      z,
      num_whole,
      num_half,
      tolerance
    };

    const response = UrlFetchApp.fetch(API_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(input),
      muteHttpExceptions: true,
    });

    const duration = (new Date() - now) / 1000;
    const timeCompleted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'h:mm:ss a');

   if (response.getResponseCode() !== 200) {
      const body = response.getContentText();
      let message;

      if (body.includes('<!DOCTYPE') || body.includes('Internal Server Error')) {
      message = '❌ API Error: Server ran out of memory or timed out. Try a smaller or different combination.';
    } else {
    message = `❌ API Error: ${body}`;
    }

    appendToCell(sheet, OUTPUT_STATUS_CELL, `\n${message}`);
    return;
  }

    const result = JSON.parse(response.getContentText());
    const summary = generateOutputText(result, input);
    sheet.getRange(OUTPUT_TEXT_CELL).setValue(summary);

    appendToCell(sheet, OUTPUT_STATUS_CELL, `\n✅ Complete in ${duration.toFixed(1)}s\nFinished at ${timeCompleted}`);

  } catch (err) {
    appendToCell(sheet, OUTPUT_STATUS_CELL, `\n❌ Error: ${err.message}`);
  }
}

function flattenInventory(data) {
  const flattened = [];
  for (let [mass, qty] of data) {
    for (let i = 0; i < qty; i++) flattened.push(parseFloat(mass));
  }
  return flattened;
}

function appendToCell(sheet, cell, text) {
  const range = sheet.getRange(cell);
  const current = range.getValue();
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, yyyy - h:mm a');
  const stampNeeded = !current.includes(ts) && (current.length === 0 || !/Running|Complete/.test(current));
  const newline = current ? '\n' : '';
  range.setValue(`${current}${newline}${stampNeeded ? `\n${text}` : text}`);
}

function generateOutputText(result, input, duration, timeCompleted) {
  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMM d, yyyy - h:mm a');

  let summary = `Date: ${timestamp}\n\n`;

  summary += `Combination Parameters:\n`;
  summary += `Total Clips (z): ${input.z}\n`;
  summary += `Whole Clip: ${input.num_whole}  Half Clip: ${input.num_half}  Tolerance: ±${input.tolerance}g\n`;
  const avgMass = result.length > 0
    ? (result.reduce((a, b) => a + b.mass, 0) / result.length).toFixed(2)
    : 'N/A';
  summary += `Avg Mass: ${avgMass}g\n\n`;

  if (result.length > 0) {
    summary += `Full Clip Combinations:\n`;
    result.forEach((clip, i) => {
      const mass = clip.mass.toFixed(1);
      const W = `[${clip.whole.map(w => w.toFixed(1)).join(' ')}]`;
      const H = `[${clip.half.map(h => h.toFixed(1)).join(' ')}]`;
      summary += `C${i + 1} M:${mass}  W:${W}  H:${H}\n`;
    });

    // Inventory summary
    const countMasses = (list) => {
      const counts = {};
      list.forEach(m => {
        const mass = parseFloat(m.toFixed(1));
        counts[mass] = (counts[mass] || 0) + 1;
      });
      return counts;
    };

    const formatCounts = (counts) => {
      return Object.entries(counts)
        .sort((a, b) => parseFloat(a[0]) - parseFloat(b[0]))
        .map(([mass, count]) => `${mass}×${count}`)
        .join('  ');
    };

    const wholeCounts = countMasses(result.flatMap(c => c.whole));
    const halfCounts = countMasses(result.flatMap(c => c.half));

    summary += `\nInventory Summary:\n`;
    summary += `Whole Clips: ${formatCounts(wholeCounts)}\n`;
    summary += `Half Clips: ${formatCounts(halfCounts)}`;
  } else {
    summary += `ERROR: No valid combination found.`;
  }

  return summary;
}


function clearStatusBox() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  sheet.getRange(OUTPUT_STATUS_CELL).clearContent();
}

function clearOutputBox() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  sheet.getRange(OUTPUT_TEXT_CELL).clearContent();
}


function removeUsedClips() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const confirmation = ui.alert("Confirm", "Are you sure you want to remove used clips from inventory?", ui.ButtonSet.YES_NO);

  if (confirmation !== ui.Button.YES) return;

  const text = sheet.getRange(OUTPUT_TEXT_CELL).getValue();
  const lines = text.split('\n');

  let wholeUsed = [], halfUsed = [];

  for (let line of lines) {
    if (line.startsWith('C')) {
      const W = line.match(/W:\[([^\]]+)\]/);
      const H = line.match(/H:\[([^\]]+)\]/);
      if (W && W[1]) {
        wholeUsed.push(...W[1].split(' ').map(Number));
      }
      if (H && H[1]) {
        halfUsed.push(...H[1].split(' ').map(Number));
      }
    }
  }

  const updateInventory = (rangeA, rangeB, used) => {
    const values = sheet.getRange(rangeA + '3:' + rangeA).getValues();
    const qtys = sheet.getRange(rangeB + '3:' + rangeB).getValues();
    let changed = false;

    for (let i = 0; i < values.length; i++) {
      const mass = parseFloat(values[i][0]);
      const qty = qtys[i][0];
      if (isNaN(mass) || qty === '') continue;
      const count = used.filter(x => Math.abs(x - mass) < 1e-3).length;
      if (count > 0 && qty >= count) {
        qtys[i][0] -= count;
        changed = true;
      }
    }

    if (changed) sheet.getRange(rangeB + '3:' + rangeB).setValues(qtys);
  };

  updateInventory('C', 'D', wholeUsed);
  updateInventory('A', 'B', halfUsed);

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, yyyy - h:mm a');
  const msg = `🛠️ Clip inventory updated at ${now}`;
  appendToCell(sheet, OUTPUT_STATUS_CELL, `\n${msg}`);
}
