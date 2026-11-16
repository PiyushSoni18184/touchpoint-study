const SPREADSHEET_ID = "1QU_62Y13pjE6SzCntliU6UkSoBNFONo1YfaqKKcbZ9o";
const SHEET_NAME = "GameTelemetry";
const HEADERS = [
  "user_id", "vision_mode", "game_mode", "trial_index", "comp_id", "comp_name",
  "comp_family", "contrast_ratio", "n_distractors", "popout_flag", "stroop_conflict",
  "seq_guidance", "fitts_A_px", "fitts_W_eff_px", "discovered_D", "matched_M_print",
  "selected_S_print", "handoff_H_seq", "selected_S_save", "false_alarms_fa",
  "sdt_outcome", "rt_print_ms", "rt_handoff_ms", "rt_total_ms", "accessibility_score",
  "sequence_success", "confidence_q", "trial_payoff_pts", "likert_seq_aid",
  "likert_stroop_hes", "sam_valence", "submitted_at"
];
const NUMERIC_FIELDS = new Set([
  "trial_index", "contrast_ratio", "n_distractors", "popout_flag", "stroop_conflict",
  "seq_guidance", "fitts_A_px", "fitts_W_eff_px", "discovered_D", "matched_M_print",
  "selected_S_print", "handoff_H_seq", "selected_S_save", "false_alarms_fa",
  "rt_print_ms", "rt_handoff_ms", "rt_total_ms", "accessibility_score",
  "sequence_success", "confidence_q", "trial_payoff_pts"
]);

function doPost(event) {
  try {
    const raw = event && event.postData && event.postData.contents;
    if (!raw || raw.length > 150000) throw new Error("Missing or oversized request.");

    const payload = JSON.parse(raw);
    const rows = payload.telemetryRows;
    const survey = payload.survey;
    if (!Array.isArray(rows) || rows.length < 1 || rows.length > 51) {
      throw new Error("Expected between 1 and 51 telemetry rows.");
    }
    validateSurvey(survey);

    const submittedAt = new Date().toISOString();
    const values = rows.map(function(row) {
      validateRow(row);
      return HEADERS.map(function(header) {
        if (header === "submitted_at") return submittedAt;
        if (header === "likert_seq_aid") return survey.l1;
        if (header === "likert_stroop_hes") return survey.l2;
        if (header === "sam_valence") return survey.sam;
        const value = row[header];
        return NUMERIC_FIELDS.has(header) ? Number(value) : safeText(value);
      });
    });

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sheet = getTelemetrySheet();
      const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
      if (firstRow.every(function(cell) { return cell === ""; })) {
        sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
        sheet.setFrozenRows(1);
      } else if (HEADERS.some(function(header, index) { return firstRow[index] !== header; })) {
        throw new Error("The GameTelemetry header row does not match the expected schema.");
      }

      sheet.getRange(sheet.getLastRow() + 1, 1, values.length, HEADERS.length).setValues(values);
    } finally {
      lock.releaseLock();
    }

    return jsonResponse({ok: true, rows: values.length});
  } catch (error) {
    return jsonResponse({ok: false, error: String(error.message || error)});
  }
}

function getTelemetrySheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
}

function validateSurvey(survey) {
  if (!survey || !isIntegerInRange(survey.sam, 1, 9) ||
      !isIntegerInRange(survey.l1, 1, 5) || !isIntegerInRange(survey.l2, 1, 5)) {
    throw new Error("Survey values are invalid.");
  }
}

function validateRow(row) {
  if (!row || typeof row !== "object" || Array.isArray(row)) throw new Error("Invalid telemetry row.");
  HEADERS.forEach(function(header) {
    if (header !== "submitted_at" && header !== "likert_seq_aid" &&
        header !== "likert_stroop_hes" && header !== "sam_valence" &&
        !Object.prototype.hasOwnProperty.call(row, header)) {
      throw new Error("A telemetry field is missing.");
    }
  });

  if (typeof row.user_id !== "string" || !row.user_id.trim() || row.user_id.length > 100 ||
      !/^C(?:0[1-9]|[1-4][0-9]|5[01])$/.test(row.comp_id) ||
      !["Normal", "Deutan_CVD", "Protan_CVD", "LowVision_Blur"].includes(row.vision_mode) ||
      !["Sprint_24", "Full_51"].includes(row.game_mode) ||
      !["Hit", "FA", "Miss"].includes(row.sdt_outcome)) {
    throw new Error("A telemetry row contains an invalid identifier or mode.");
  }

  NUMERIC_FIELDS.forEach(function(field) {
    const rawValue = row[field];
    if ((typeof rawValue !== "number" && typeof rawValue !== "string") ||
        (typeof rawValue === "string" && !rawValue.trim())) {
      throw new Error("A numeric telemetry field is invalid.");
    }
    const value = Number(rawValue);
    if (!Number.isFinite(value)) throw new Error("A numeric telemetry field is invalid.");
    if (["popout_flag", "stroop_conflict", "discovered_D", "matched_M_print",
         "selected_S_print", "handoff_H_seq", "selected_S_save", "sequence_success"].includes(field) &&
        value !== 0 && value !== 1) {
      throw new Error("A telemetry flag is invalid.");
    }
  });
  if (!isIntegerInRange(Number(row.trial_index), 1, 51)) throw new Error("Trial index is invalid.");
}

function isIntegerInRange(value, min, max) {
  return Number.isInteger(Number(value)) && Number(value) >= min && Number(value) <= max;
}

function safeText(value) {
  if (typeof value !== "string" || value.length > 200) throw new Error("A text telemetry field is invalid.");
  return /^[\s]*[=+\-@]/.test(value) ? "'" + value : value;
}

function jsonResponse(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
