/**
 * The Squad Health Check spreadsheet name.
 * Used to prefix the names of survey forms and response sheets.
 */
const SQUAD_HEALTH_CHECK_SHEET_PREFIX = "Squad Health Check"

/**
 * The titles for the perception and trend dimension measurements
 * (and thus graph sheet prefixes).
 */
const PERCEPTION_TITLE = "Perception"
const TREND_TITLE = "Trend"

/**
 * An array to map sheet column interger indexes to letters.
 * Covers 1 (A) through 52 (AZ),
 * plus a dummy at 0 because of 0-index vs. 1-index BS.
 */
const INTEGERS_TO_COLUMNS = Object.freeze(
  Array
    .from(
      { length: 27 }, (_, i) =>
        String.fromCharCode('A'.charCodeAt(0) - 1 + i)
    )
    .concat(
      Array.from(
        { length: 26 }, (_, i) =>
          `A${String.fromCharCode('A'.charCodeAt(0) + i)}`
      )
    )
)
