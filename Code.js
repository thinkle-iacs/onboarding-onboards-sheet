const HEADER_SEARCH_ROWS = 3;
const ONBOARDING_BASE_URL = "https://tinyurl.com/iacs-onboard";

const FIELDS = {
  iacsEmailNeeded: "iacsEmailNeeded",
  googleOnboardingLink: "googleOnboardingLink",
  lastFirst: "lastFirst",
  firstName: "firstName",
  lastName: "lastName",
  iacsUsername: "iacsUsername",
  ou: "ou",
  position: "position",
  school: "school",
  department: "department",
};

const HEADER_MATCHERS = {
  [FIELDS.iacsEmailNeeded]: [/^iacs.*email.*needed$/i],
  [FIELDS.googleOnboardingLink]: [/google.*onboard.*link/i, /google.*link/i],
  [FIELDS.lastFirst]: [/^last\s*,\s*first$/i],
  [FIELDS.firstName]: [/^first(\s*name)?$/i],
  [FIELDS.lastName]: [/^last(\s*name)?$/i],
  [FIELDS.iacsUsername]: [/^iacs.*(user|account|email)/i],
  [FIELDS.ou]: [/^ou$/i, /org.*unit/i],
  [FIELDS.position]: [/position|title/i],
  [FIELDS.school]: [/school|organization/i],
  [FIELDS.department]: [/department|dept/i],
};

const TAB_DEFAULT_RULES = [
  {
    matcher: /sub/i,
    ou: "/Staff/Teachers/Substitutes",
  },
  {
    matcher: /coach|athlet/i,
    ou: "/Staff/Athletics",
  },
];

function onEdit(e) {
  handleIacsEmailNeededEdit(e);
}

function handleIacsEmailNeededEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const headerInfo = getHeaderInfo_(sheet);
  const headerMap = headerInfo.headerMap;
  const neededColumn = headerMap[FIELDS.iacsEmailNeeded];
  const linkColumn = headerMap[FIELDS.googleOnboardingLink];

  if (!neededColumn || !linkColumn) return;
  if (!rangeIncludesColumn_(e.range, neededColumn)) return;

  const startRow = Math.max(e.range.getRow(), headerInfo.headerRow + 1);
  const endRow = e.range.getLastRow();
  for (let row = startRow; row <= endRow; row++) {
    const editedValue = sheet.getRange(row, neededColumn).getValue();
    if (!isYes_(editedValue)) continue;

    const namedValues = getNamedValuesForRow_(sheet, headerMap, row);
    const tabDefaults = getTabDefaults_(sheet.getName());
    applyRowDefaults_(sheet, headerMap, row, namedValues, tabDefaults);

    sheet
      .getRange(row, linkColumn)
      .setValue(buildGoogleOnboardingLink_(namedValues));
  }
}

function getHeaderInfo_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return { headerRow: 0, headerMap: {} };

  const rowCount = Math.min(HEADER_SEARCH_ROWS, sheet.getLastRow());
  if (rowCount < 1) return { headerRow: 0, headerMap: {} };

  const rows = sheet.getRange(1, 1, rowCount, lastColumn).getValues();
  let best = { headerRow: 0, headerMap: {}, score: 0 };

  rows.forEach((headers, index) => {
    const headerMap = buildHeaderMap_(headers);
    const score = Object.keys(headerMap).length;
    const hasRequiredHeaders =
      headerMap[FIELDS.iacsEmailNeeded] &&
      headerMap[FIELDS.googleOnboardingLink];

    if (hasRequiredHeaders && score > best.score) {
      best = { headerRow: index + 1, headerMap, score };
    }
  });

  return { headerRow: best.headerRow, headerMap: best.headerMap };
}

function buildHeaderMap_(headers) {
  return headers.reduce((map, header, index) => {
    const field = getFieldForHeader_(header);
    if (field && !map[field]) map[field] = index + 1;
    return map;
  }, {});
}

function getNamedValuesForRow_(sheet, headerMap, row) {
  const namedValues = {};
  Object.keys(headerMap).forEach((field) => {
    namedValues[field] = sheet.getRange(row, headerMap[field]).getValue();
  });

  const parsedName = parseLastFirst_(namedValues[FIELDS.lastFirst]);
  if (!namedValues[FIELDS.firstName] && parsedName.firstName) {
    namedValues[FIELDS.firstName] = parsedName.firstName;
  }
  if (!namedValues[FIELDS.lastName] && parsedName.lastName) {
    namedValues[FIELDS.lastName] = parsedName.lastName;
  }

  return namedValues;
}

function applyRowDefaults_(sheet, headerMap, row, namedValues, defaults) {
  if (!defaults.ou || namedValues[FIELDS.ou]) return;

  namedValues[FIELDS.ou] = defaults.ou;
  if (headerMap[FIELDS.ou]) {
    sheet.getRange(row, headerMap[FIELDS.ou]).setValue(defaults.ou);
  }
}

function getTabDefaults_(sheetName) {
  const matchingRule = TAB_DEFAULT_RULES.find((rule) =>
    rule.matcher.test(sheetName)
  );
  return matchingRule || {};
}

function buildGoogleOnboardingLink_(namedValues) {
  const iacsAccount = String(namedValues[FIELDS.iacsUsername] || "").trim();
  const firstName = String(namedValues[FIELDS.firstName] || "").trim();
  const lastName = String(namedValues[FIELDS.lastName] || "").trim();
  const username =
    iacsAccount.split("@")[0] ||
    (firstName && lastName ? sanitizeUsername_(firstName[0] + lastName) : "");

  const queryParams = {
    u: username,
    first: firstName,
    last: lastName,
    organization: getOnboardingOrganization_(namedValues[FIELDS.school]),
    title: namedValues[FIELDS.position],
    department: namedValues[FIELDS.department],
    ou: namedValues[FIELDS.ou],
  };

  const queryString = Object.keys(queryParams)
    .filter((key) => queryParams[key])
    .map(
      (key) =>
        `${encodeURIComponent(key)}=${encodeURIComponent(queryParams[key])}`
    )
    .join("&");

  return queryString
    ? `${ONBOARDING_BASE_URL}?${queryString}`
    : ONBOARDING_BASE_URL;
}

function parseLastFirst_(value) {
  const text = String(value || "").trim();
  if (!text) return {};

  const parts = text.split(",");
  if (parts.length >= 2) {
    return {
      lastName: parts[0].trim(),
      firstName: parts
        .slice(1)
        .join(",")
        .trim()
        .split(/\s+/)[0],
    };
  }

  const nameParts = text.split(/\s+/);
  if (nameParts.length < 2) return { firstName: text };

  return {
    firstName: nameParts[0],
    lastName: nameParts.slice(1).join(" "),
  };
}

function getOnboardingOrganization_(school) {
  const schoolText = String(school || "").trim();
  const normalizedSchool = schoolText.toUpperCase();
  if (["HS", "MS", "IACS"].indexOf(normalizedSchool) > -1) {
    return normalizedSchool;
  }
  if (normalizedSchool.indexOf("HIGH") > -1) return "HS";
  if (normalizedSchool.indexOf("MIDDLE") > -1) return "MS";
  return schoolText;
}

function sanitizeUsername_(name) {
  return String(name || "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[^a-z0-9._%+-]+/g, "")
    .replace(/[.]/g, "");
}

function isYes_(value) {
  const text = String(value || "")
    .trim()
    .toLowerCase();
  return text === "yes" || text === "y";
}

function rangeIncludesColumn_(range, column) {
  return range.getColumn() <= column && range.getLastColumn() >= column;
}

function getFieldForHeader_(header) {
  const text = String(header || "").trim();
  if (!text) return "";

  for (const field in HEADER_MATCHERS) {
    if (HEADER_MATCHERS[field].some((matcher) => matcher.test(text))) {
      return field;
    }
  }
  return "";
}
