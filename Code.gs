const SPREADSHEET_ID = "1QU_62Y13pjE6SzCntliU6UkSoBNFONo1YfaqKKcbZ9o";
const SHEET_NAME = "GameTelemetry";
const HEADERS = [
  "user_id", "vision_mode", "game_mode", "trial_index", "comp_id", "comp_name",
  "comp_family", "contrast_ratio", "n_distractors", "popout_flag", "stroop_conflict",
  "seq_guidance", "fitts_A_px", "fitts_W_eff_px", "discovered_D", "matched_M_print",
  "selected_S_print", "handoff_H_seq", "selected_S_save", "false_alarms_fa",
  "sdt_outcome", "rt_print_ms", "rt_handoff_ms", "rt_total_ms", "accessibility_score",
  "sequence_success", "confidence_q", "trial_payoff_pts", "likert_seq_aid",
  "likert_stroop_hes", "sam_valence", "submitted_at", "trial_status", "skipped",
  "confidence_label"
];
const NUMERIC_FIELDS = new Set([
  "trial_index", "contrast_ratio", "n_distractors", "popout_flag", "stroop_conflict",
  "seq_guidance", "fitts_A_px", "fitts_W_eff_px", "discovered_D", "matched_M_print",
  "selected_S_print", "handoff_H_seq", "selected_S_save", "false_alarms_fa",
  "rt_print_ms", "rt_handoff_ms", "rt_total_ms", "accessibility_score",
  "sequence_success", "confidence_q", "trial_payoff_pts", "skipped"
]);
const OPTIONAL_NUMERIC_FIELDS = new Set([
  "fitts_A_px", "fitts_W_eff_px", "rt_print_ms", "rt_handoff_ms",
  "accessibility_score", "confidence_q"
]);
const BASE_HEADER_COUNT = HEADERS.length - 3;
const CONFIDENCE_LABELS = ["Sure", "Somewhat Sure", "Not Sure", "Skipped"];
const SURVEY_LABELS = ["Sure", "Somewhat Sure", "Not Sure"];

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
        const value = row[header] === undefined
          ? (header === "trial_status" ? "Completed" : header === "skipped" ? 0 : "")
          : row[header];
        if (NUMERIC_FIELDS.has(header)) {
          return OPTIONAL_NUMERIC_FIELDS.has(header) && value === "NA" ? "NA" : Number(value);
        }
        return value === "" ? "" : safeText(value);
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
      } else if (HEADERS.slice(0, BASE_HEADER_COUNT).some(function(header, index) {
        return firstRow[index] !== header;
      })) {
        throw new Error("The GameTelemetry header row does not match the expected schema.");
      } else {
        for (let index = BASE_HEADER_COUNT; index < HEADERS.length; index++) {
          if (firstRow[index] === "") {
            sheet.getRange(1, index + 1).setValue(HEADERS[index]);
          } else if (firstRow[index] !== HEADERS[index]) {
            throw new Error("The GameTelemetry header row does not match the expected schema.");
          }
        }
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
      !SURVEY_LABELS.includes(survey.l1) || !SURVEY_LABELS.includes(survey.l2)) {
    throw new Error("Survey values are invalid.");
  }
}

function validateRow(row) {
  if (!row || typeof row !== "object" || Array.isArray(row)) throw new Error("Invalid telemetry row.");
  HEADERS.slice(0, BASE_HEADER_COUNT).forEach(function(header) {
    if (header !== "submitted_at" && header !== "likert_seq_aid" &&
        header !== "likert_stroop_hes" && header !== "sam_valence" &&
        !Object.prototype.hasOwnProperty.call(row, header)) {
      throw new Error("A telemetry field is missing.");
    }
  });
  if (row.trial_status !== undefined && !["Completed", "Skipped"].includes(row.trial_status)) {
    throw new Error("Trial status is invalid.");
  }
  if (row.skipped !== undefined && row.skipped !== 0 && row.skipped !== 1) {
    throw new Error("Skip flag is invalid.");
  }
  if (row.confidence_label !== undefined && !CONFIDENCE_LABELS.includes(row.confidence_label)) {
    throw new Error("Confidence label is invalid.");
  }

  if (typeof row.user_id !== "string" || !row.user_id.trim() || row.user_id.length > 100 ||
      !/^C(?:0[1-9]|[1-4][0-9]|5[01])$/.test(row.comp_id) ||
      !["Normal", "Deutan_CVD", "Protan_CVD", "LowVision_Blur"].includes(row.vision_mode) ||
      !["Sprint_24", "Full_51"].includes(row.game_mode) ||
      !["Hit", "FA", "Miss", "Skip"].includes(row.sdt_outcome)) {
    throw new Error("A telemetry row contains an invalid identifier or mode.");
  }

  if ((row.trial_status === "Skipped") !== (row.sdt_outcome === "Skip") ||
      (row.skipped !== undefined && (row.skipped === 1) !== (row.sdt_outcome === "Skip"))) {
    throw new Error("Trial status and skip flag do not match.");
  }

  NUMERIC_FIELDS.forEach(function(field) {
    if (field === "skipped" && row.skipped === undefined) return;
    if (OPTIONAL_NUMERIC_FIELDS.has(field) && row[field] === "NA") return;
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
