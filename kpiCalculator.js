function safeNumber(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  const numeric = Number(String(value).replace(/[,%]/g, "").trim());
  return Number.isFinite(numeric) ? numeric : null;
}

function titleCaseMonth(date) {
  return date.toLocaleString("en-US", { month: "long", year: "numeric" });
}

function parseWeekEnding(value) {
  if (!value) return null;
  const [month, day, year] = String(value).trim().split("/");
  const parsed = new Date(Number(year), Number(month) - 1, Number(day));
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function parseDateRange(value) {
  if (!value) return null;
  const parts = String(value)
    .split("-")
    .map((part) => part.trim())
    .filter(Boolean);
  const endValue = parts.at(-1) || String(value).trim();
  return parseWeekEnding(endValue);
}

function parseTimestamp(value) {
  if (!value) return null;
  const parsed = new Date(String(value).trim());
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function parseSourceDate(value) {
  if (!value) return null;
  const trimmed = String(value).trim();
  if (!trimmed) return null;

  const direct = new Date(trimmed);
  if (!Number.isNaN(direct.getTime())) {
    return new Date(direct.getFullYear(), direct.getMonth(), direct.getDate());
  }

  const parts = trimmed.split(/[/-]/).map((part) => Number(part));
  if (parts.length === 3 && parts.every((part) => Number.isFinite(part))) {
    const [first, second, third] = parts;
    const parsed = first > 31
      ? new Date(first, second - 1, third)
      : new Date(third, first - 1, second);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  return null;
}

function formatLastUpdatedDisplay(rawValue, parsedValue = null) {
  const trimmed = String(rawValue ?? "").trim();
  if (trimmed) return `${trimmed} EST`;
  if (parsedValue instanceof Date && !Number.isNaN(parsedValue.getTime())) {
    return `${parsedValue.toLocaleString("en-US")} EST`;
  }
  return "N/A";
}

function secondsFromDuration(duration) {
  if (!duration) return 0;

  const parts = String(duration)
    .trim()
    .split(":")
    .map((part) => Number(part));

  if (parts.some((part) => Number.isNaN(part))) return 0;
  if (parts.length === 3) {
    const [hours, minutes, seconds] = parts;
    return hours * 3600 + minutes * 60 + seconds;
  }
  if (parts.length === 2) {
    const [minutes, seconds] = parts;
    return minutes * 60 + seconds;
  }
  return parts[0] ?? 0;
}

function formatDurationFromSeconds(totalSeconds) {
  if (totalSeconds === null || totalSeconds === undefined || Number.isNaN(totalSeconds)) return "";
  const normalized = Math.max(0, Math.round(totalSeconds));
  const hours = Math.floor(normalized / 3600);
  const minutes = Math.floor((normalized % 3600) / 60);
  const seconds = normalized % 60;
  if (hours > 0) {
    return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
  }
  return `${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
}

function normalizeEmail(value) {
  return String(value ?? "").trim().toLowerCase();
}

function normalizeDisplayName(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function normalizeAlias(value) {
  return normalizeDisplayName(value)
    .toLowerCase()
    .replace(/phillies$/i, "")
    .replace(/,/g, "")
    .replace(/\./g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function makeKey(identity, weekEnding) {
  return `${String(identity ?? "").trim()}__${String(weekEnding ?? "").trim()}`;
}

export function calculateTransferRate(transferCount, firstTimeCaller) {
  if (transferCount === null || firstTimeCaller === null || !firstTimeCaller) return null;
  return transferCount / firstTimeCaller;
}

export function calculateTransferScore(transferRate) {
  if (transferRate === null || transferRate === undefined || Number.isNaN(transferRate)) return null;
  const percent = transferRate * 100;
  if (percent < 4) return 1;
  if (percent <= 8) return 2;
  if (percent <= 10) return 3;
  if (percent <= 15) return 4;
  return 5;
}

export function calculateAdmitsScore(count) {
  if (count === null || count === undefined || Number.isNaN(count)) return null;
  if (!count) return 1;
  if (count <= 4) return 2;
  if (count <= 8) return 3;
  if (count <= 13) return 4;
  return 5;
}

export function calculateAHT(inboundMinutes, holdTime, inboundCalls) {
  if (inboundCalls === null || inboundCalls === undefined || !inboundCalls) return null;
  if (!inboundMinutes && !holdTime) return null;
  const totalSeconds = secondsFromDuration(inboundMinutes) + secondsFromDuration(holdTime);
  return totalSeconds / inboundCalls;
}

export function calculateAHTScore(ahtSeconds) {
  if (ahtSeconds === null || ahtSeconds === undefined || Number.isNaN(ahtSeconds)) return null;
  if (ahtSeconds < 104) return 5;
  if (ahtSeconds <= 140) return 4;
  if (ahtSeconds <= 176) return 3;
  if (ahtSeconds <= 212) return 2;
  return 1;
}

export function calculateAttendanceScore(attendancePercentRaw) {
  if (attendancePercentRaw === null || attendancePercentRaw === undefined) {
    return null;
  }
  if (attendancePercentRaw === "") {
    return 4;
  }
  const attendancePercent = safeNumber(attendancePercentRaw);
  if (attendancePercent === null) return null;
  if (attendancePercent >= 100) return 5;
  if (attendancePercent >= 95) return 3;
  if (attendancePercent >= 90) return 2;
  return 1;
}

export function calculateQAScore(qaPercentRaw) {
  const qaPercent = safeNumber(qaPercentRaw);
  if (qaPercent === null) return null;
  if (qaPercent >= 100) return 5;
  if (qaPercent >= 99) return 4;
  if (qaPercent >= 98) return 3;
  if (qaPercent >= 95) return 2;
  return 1;
}

export function calculatePerformanceScore(transferScore, admitsScore, ahtScore) {
  const scores = [transferScore, admitsScore, ahtScore]
    .filter((value) => typeof value === "number" && !Number.isNaN(value));
  if (!scores.length) return null;
  return scores.reduce((sum, value) => sum + value, 0) / scores.length;
}

export function getOverallComposition(performanceScore, attendanceScore, qaScore) {
  const baseWeights = {
    performance: 0.5,
    attendance: 0.25,
    qa: 0.25,
  };
  const availability = {
    performance: typeof performanceScore === "number" && !Number.isNaN(performanceScore),
    attendance: typeof attendanceScore === "number" && !Number.isNaN(attendanceScore),
    qa: typeof qaScore === "number" && !Number.isNaN(qaScore),
  };
  const activeWeight = Object.entries(baseWeights).reduce(
    (sum, [key, weight]) => sum + (availability[key] ? weight : 0),
    0
  );

  return {
    activeWeight,
    includesQa: availability.qa,
    weights: {
      performance: availability.performance && activeWeight ? baseWeights.performance / activeWeight : 0,
      attendance: availability.attendance && activeWeight ? baseWeights.attendance / activeWeight : 0,
      qa: availability.qa && activeWeight ? baseWeights.qa / activeWeight : 0,
    },
  };
}

export function calculateOverallScore(performanceScore, attendanceScore, qaScore) {
  const composition = getOverallComposition(performanceScore, attendanceScore, qaScore);
  if (!composition.activeWeight) return null;

  return (
    performanceScore * composition.weights.performance +
    attendanceScore * composition.weights.attendance +
    qaScore * composition.weights.qa
  );
}

function formatAHT(seconds) {
  if (seconds === null || seconds === undefined || Number.isNaN(seconds)) {
    return "N/A";
  }
  const mins = Math.floor(seconds / 60);
  const secs = Math.round(seconds % 60);
  return `${String(mins).padStart(2, "0")}:${String(secs).padStart(2, "0")}`;
}

function formatPercent(value, digits = 2) {
  if (value === null || value === undefined || Number.isNaN(value)) {
    return "N/A";
  }
  return `${Number(value).toFixed(digits)}%`;
}

function buildPrimaryKeyLookup(rows) {
  const byEmail = new Map();
  const byAlias = new Map();

  rows.forEach((row) => {
    const email = normalizeEmail(row.email);
    const displayName = normalizeDisplayName(row.displayName);
    const aliases = [
      displayName,
      normalizeDisplayName(row.altNameOne),
      normalizeDisplayName(row.altNameTwo),
    ].filter(Boolean);

    const entry = {
      identity: displayName || email,
      email,
      displayName,
      aliases,
    };

    if (email) {
      const emailEntries = byEmail.get(email) || [];
      emailEntries.push(entry);
      byEmail.set(email, emailEntries);
    }

    aliases.forEach((alias) => {
      const normalized = normalizeAlias(alias);
      if (!normalized) return;
      const aliasEntries = byAlias.get(normalized) || [];
      aliasEntries.push(entry);
      byAlias.set(normalized, aliasEntries);
    });
  });

  return { byEmail, byAlias };
}

function resolvePerson(primaryKeyLookup, email, fallbackName) {
  const normalizedEmail = normalizeEmail(email);
  const normalizedFallback = normalizeAlias(fallbackName);
  const emailMatches = primaryKeyLookup.byEmail.get(normalizedEmail) || [];

  if (emailMatches.length === 1) {
    return emailMatches[0];
  }

  if (emailMatches.length > 1 && normalizedFallback) {
    const exact = emailMatches.find((entry) =>
      entry.aliases.some((alias) => normalizeAlias(alias) === normalizedFallback)
    );
    if (exact) return exact;
  }

  if (normalizedFallback) {
    const aliasMatches = primaryKeyLookup.byAlias.get(normalizedFallback) || [];
    if (aliasMatches.length === 1) return aliasMatches[0];
    if (aliasMatches.length > 1) {
      const emailExact = aliasMatches.find((entry) => entry.email === normalizedEmail);
      if (emailExact) return emailExact;
      return aliasMatches[0];
    }
  }

  return {
    identity: normalizedEmail || normalizeDisplayName(fallbackName),
    email: normalizedEmail,
    displayName: normalizeDisplayName(fallbackName),
    aliases: [normalizeDisplayName(fallbackName)].filter(Boolean),
  };
}

function aggregateRealtimeRows(rows, primaryKeyLookup) {
  const registry = new Map();

  rows.forEach((row) => {
    const dateRange = String(row.dateRange || row.date || row.weekEnding || "").trim();
    const person = resolvePerson(primaryKeyLookup, row.email, row.agentName);
    const aggregateIdentity = normalizeEmail(person.email) || person.identity || normalizeDisplayName(row.agentName);
    if (!dateRange || !aggregateIdentity) return;

    const key = `${dateRange}__${aggregateIdentity}`;
    const existing = registry.get(key) || {
      ...row,
      dateRange,
      date: row.date || "",
      agentName: person.displayName || normalizeDisplayName(row.agentName),
      email: person.email || normalizeEmail(row.email),
      firstTimeCallerTotal: 0,
      transferCountTotal: 0,
      inboundCallsTotal: 0,
      inboundMinutesTotalSeconds: 0,
      holdTimeTotalSeconds: 0,
      lastUpdatedAt: null,
      lastUpdated: "",
    };

    existing.firstTimeCallerTotal += safeNumber(row.firstTimeCaller) || 0;
    existing.transferCountTotal += safeNumber(row.transferCount) || 0;
    existing.inboundCallsTotal += safeNumber(row.inboundCalls) || 0;
    existing.inboundMinutesTotalSeconds += secondsFromDuration(row.inboundMinutes);
    existing.holdTimeTotalSeconds += secondsFromDuration(row.holdTime);

    const parsedLastUpdated = parseTimestamp(row.lastUpdated);
    if (
      parsedLastUpdated instanceof Date &&
      !Number.isNaN(parsedLastUpdated.getTime()) &&
      (!existing.lastUpdatedAt || parsedLastUpdated.getTime() >= existing.lastUpdatedAt.getTime())
    ) {
      existing.lastUpdatedAt = parsedLastUpdated;
      existing.lastUpdated = row.lastUpdated || existing.lastUpdated;
    }

    registry.set(key, existing);
  });

  return [...registry.values()].map((row) => ({
    ...row,
    firstTimeCaller: row.firstTimeCallerTotal ? String(row.firstTimeCallerTotal) : "",
    transferCount: row.transferCountTotal ? String(row.transferCountTotal) : "",
    inboundCalls: row.inboundCallsTotal ? String(row.inboundCallsTotal) : "",
    inboundMinutes: formatDurationFromSeconds(row.inboundMinutesTotalSeconds),
    holdTime: formatDurationFromSeconds(row.holdTimeTotalSeconds),
    lastUpdated: row.lastUpdated || (row.lastUpdatedAt ? row.lastUpdatedAt.toLocaleString("en-US") : ""),
  }));
}

function aggregateRealtimeDailyRows(rows, primaryKeyLookup) {
  const registry = new Map();

  rows.forEach((row) => {
    const dateAt = parseSourceDate(row.date);
    const person = resolvePerson(primaryKeyLookup, row.email, row.agentName);
    const aggregateIdentity = normalizeEmail(person.email) || person.identity || normalizeDisplayName(row.agentName);
    if (!dateAt || !aggregateIdentity) return;

    const dateKey = `${dateAt.getFullYear()}-${String(dateAt.getMonth() + 1).padStart(2, "0")}-${String(dateAt.getDate()).padStart(2, "0")}`;
    const key = `${dateKey}__${aggregateIdentity}`;
    const existing = registry.get(key) || {
      ...row,
      date: row.date || dateKey,
      dateAt,
      dateRange: row.dateRange || "",
      agentName: person.displayName || normalizeDisplayName(row.agentName),
      email: person.email || normalizeEmail(row.email),
      firstTimeCallerTotal: 0,
      transferCountTotal: 0,
      inboundCallsTotal: 0,
      inboundMinutesTotalSeconds: 0,
      holdTimeTotalSeconds: 0,
      lastUpdatedAt: null,
      lastUpdated: "",
    };

    existing.firstTimeCallerTotal += safeNumber(row.firstTimeCaller) || 0;
    existing.transferCountTotal += safeNumber(row.transferCount) || 0;
    existing.inboundCallsTotal += safeNumber(row.inboundCalls) || 0;
    existing.inboundMinutesTotalSeconds += secondsFromDuration(row.inboundMinutes);
    existing.holdTimeTotalSeconds += secondsFromDuration(row.holdTime);

    const parsedLastUpdated = parseTimestamp(row.lastUpdated);
    if (
      parsedLastUpdated instanceof Date &&
      !Number.isNaN(parsedLastUpdated.getTime()) &&
      (!existing.lastUpdatedAt || parsedLastUpdated.getTime() >= existing.lastUpdatedAt.getTime())
    ) {
      existing.lastUpdatedAt = parsedLastUpdated;
      existing.lastUpdated = row.lastUpdated || existing.lastUpdated;
    }

    registry.set(key, existing);
  });

  return [...registry.values()].map((row) => {
    const firstTimeCaller = row.firstTimeCallerTotal || 0;
    const transferCount = row.transferCountTotal || 0;
    const inboundCalls = row.inboundCallsTotal || 0;
    const inboundMinutesSeconds = row.inboundMinutesTotalSeconds || 0;
    const holdTimeSeconds = row.holdTimeTotalSeconds || 0;
      const transferRate = calculateTransferRate(
        transferCount === null || transferCount === undefined ? null : transferCount,
        firstTimeCaller === null || firstTimeCaller === undefined ? null : firstTimeCaller
      );
    const ahtSeconds = inboundCalls > 0 ? (inboundMinutesSeconds + holdTimeSeconds) / inboundCalls : null;

    return {
      key: `${normalizeEmail(row.email) || row.agentName}__${row.date}`,
      identity: normalizeEmail(row.email) || row.agentName,
      email: normalizeEmail(row.email),
      agentName: row.agentName,
      date: row.date,
      dateAt: row.dateAt,
      dateRange: row.dateRange || "",
      monthLabel: titleCaseMonth(row.dateAt),
      firstTimeCaller,
      transferCount,
      inboundCalls,
      inboundMinutesSeconds,
      inboundMinutes: formatDurationFromSeconds(inboundMinutesSeconds),
      holdTimeSeconds,
      holdTime: formatDurationFromSeconds(holdTimeSeconds),
      transferRate,
      transferRatePercent: transferRate === null ? null : transferRate * 100,
      transferRateDisplay: formatPercent(transferRate === null ? null : transferRate * 100),
      transferScore: calculateTransferScore(transferRate),
      ahtSeconds,
      ahtDisplay: formatAHT(ahtSeconds),
      ahtScore: calculateAHTScore(ahtSeconds),
      lastUpdated: row.lastUpdated || "",
      lastUpdatedAt: row.lastUpdatedAt,
      lastUpdatedDisplay: formatLastUpdatedDisplay(row.lastUpdated, row.lastUpdatedAt),
    };
  });
}

function mergeRows(rows, registry, sourceName, primaryKeyLookup) {
  rows.forEach((row) => {
    const weekEnding = String(row.dateRange || row.date || row.weekEnding || "").trim();
    const email = row.email;
    const person = resolvePerson(primaryKeyLookup, email, row.agentName);
    const key = makeKey(person.identity, weekEnding);

    if (!weekEnding || !person.identity) return;

    const existing = registry.get(key) || {
      key,
      identity: person.identity,
      weekEnding,
      weekDate: parseDateRange(weekEnding),
      monthLabel: "",
      email: person.email || normalizeEmail(email),
      dateRange: weekEnding,
      agentName: person.displayName || "",
      sourceFlags: {},
    };

    Object.assign(existing, row);
    existing.identity = person.identity;
    existing.agentName = person.displayName || normalizeDisplayName(row.agentName || existing.agentName);
    existing.email = person.email || normalizeEmail(email);
    existing.weekEnding = existing.weekEnding || weekEnding;
    existing.weekDate = existing.weekDate || parseDateRange(weekEnding);
    existing.dateRange = existing.dateRange || weekEnding;
    existing.monthLabel = existing.weekDate ? titleCaseMonth(existing.weekDate) : "Unknown";
    existing.sourceFlags[sourceName] = true;
    if (row.lastUpdated) {
      existing[`${sourceName}LastUpdated`] = row.lastUpdated;
    }
    registry.set(key, existing);
  });
}

function average(items, selector) {
  const values = items
    .map((item) => selector(item))
    .filter((value) => typeof value === "number" && !Number.isNaN(value));
  if (!values.length) return null;
  const total = values.reduce((sum, value) => sum + value, 0);
  return total / values.length;
}

export function groupByWeek(records) {
  return records.reduce((map, record) => {
    const bucket = map.get(record.weekEnding) || [];
    bucket.push(record);
    map.set(record.weekEnding, bucket);
    return map;
  }, new Map());
}

export function getTopPerformers(records, limit = 5) {
  return [...records]
    .filter((record) => typeof record.overallScore === "number" && !Number.isNaN(record.overallScore))
    .sort((left, right) => right.overallScore - left.overallScore)
    .slice(0, limit);
}

export function getBottomPerformers(records, limit = 5) {
  return [...records]
    .filter((record) => typeof record.overallScore === "number" && !Number.isNaN(record.overallScore))
    .sort((left, right) => left.overallScore - right.overallScore)
    .slice(0, limit);
}

export function buildKpiDataset(rawDatasets) {
  const registry = new Map();
  const primaryKeyLookup = buildPrimaryKeyLookup(rawDatasets.primaryKey || []);
  const aggregatedRealtimeRows = aggregateRealtimeRows(rawDatasets.realtime || [], primaryKeyLookup);

  mergeRows(aggregatedRealtimeRows, registry, "realtime", primaryKeyLookup);
  ["admits", "qa", "attendance"].forEach((sourceName) => {
    mergeRows(rawDatasets[sourceName] || [], registry, sourceName, primaryKeyLookup);
  });

  const records = [...registry.values()]
    .map((record) => {
      const admitsCount = safeNumber(record.admitsCount);
      const firstTimeCaller = safeNumber(record.firstTimeCaller);
      const transferCount = safeNumber(record.transferCount);
      const inboundCalls = safeNumber(record.inboundCalls);
      const inboundMinutesSeconds = secondsFromDuration(record.inboundMinutes);
      const holdTimeSeconds = secondsFromDuration(record.holdTime);
      const rawAdmitsCount = safeNumber(record.admitsCount);
      const transferRate = calculateTransferRate(transferCount, firstTimeCaller);
      const transferScore = calculateTransferScore(transferRate);
      const admitsScore = calculateAdmitsScore(rawAdmitsCount);
      const ahtSeconds = calculateAHT(record.inboundMinutes, record.holdTime, inboundCalls);
      const ahtScore = calculateAHTScore(ahtSeconds);
      const attendanceScore = calculateAttendanceScore(record.attendancePercent);
      const qaScore = calculateQAScore(record.qaPercent);
      const performanceScore = calculatePerformanceScore(transferScore, admitsScore, ahtScore);
        const overallComposition = getOverallComposition(performanceScore, attendanceScore, qaScore);
        const overallScore = calculateOverallScore(performanceScore, attendanceScore, qaScore);
        const admitsLastUpdatedAt = parseTimestamp(record.admitsLastUpdated);
        const attendanceLastUpdatedAt = parseTimestamp(record.attendanceLastUpdated);
        const qaLastUpdatedAt = parseTimestamp(record.qaLastUpdated);

      return {
        ...record,
        admitsCount,
        firstTimeCaller,
        transferCount,
        inboundCalls,
        inboundMinutesSeconds,
        holdTimeSeconds,
        transferRate,
        transferRatePercent: transferRate === null ? null : transferRate * 100,
        transferRateDisplay: formatPercent(transferRate === null ? null : transferRate * 100),
        transferScore,
        admitsScore,
        ahtSeconds,
        ahtDisplay: formatAHT(ahtSeconds),
        ahtScore,
        attendancePercentValue: safeNumber(record.attendancePercent),
        attendancePercentDisplay: formatPercent(safeNumber(record.attendancePercent)),
        attendanceScore,
        qaPercentValue: safeNumber(record.qaPercent),
        qaPercentDisplay: formatPercent(safeNumber(record.qaPercent)),
        qaScore,
        performanceScore,
        overallScore,
        overallIncludesQa: overallComposition.includesQa,
          overallWeights: overallComposition.weights,
          admitsLastUpdatedAt,
          admitsLastUpdatedDisplay: formatLastUpdatedDisplay(record.admitsLastUpdated, admitsLastUpdatedAt),
          attendanceLastUpdatedAt,
          attendanceLastUpdatedDisplay: formatLastUpdatedDisplay(record.attendanceLastUpdated, attendanceLastUpdatedAt),
          qaLastUpdatedAt,
          qaLastUpdatedDisplay: formatLastUpdatedDisplay(record.qaLastUpdated, qaLastUpdatedAt),
        };
    })
    .sort((left, right) => {
      const leftTime = left.weekDate?.getTime() ?? 0;
      const rightTime = right.weekDate?.getTime() ?? 0;
      return leftTime - rightTime || left.agentName.localeCompare(right.agentName);
    });

  const weekGroups = groupByWeek(records);
  const weeklyAverages = [...weekGroups.entries()]
    .map(([weekEnding, items]) => ({
      weekEnding,
      weekDate: items[0]?.weekDate || parseWeekEnding(weekEnding),
      monthLabel: items[0]?.monthLabel || "Unknown",
      transferScore: average(items, (item) => item.transferScore),
      admitsScore: average(items, (item) => item.admitsScore),
      ahtScore: average(items, (item) => item.ahtScore),
      attendanceScore: average(items, (item) => item.attendanceScore),
      qaScore: average(items, (item) => item.qaScore),
      performanceScore: average(items, (item) => item.performanceScore),
      overallScore: average(items, (item) => item.overallScore),
      overallIncludesQa: items.some((item) => item.overallIncludesQa),
      overallWeights: {
        performance: average(items, (item) => item.overallWeights?.performance),
        attendance: average(items, (item) => item.overallWeights?.attendance),
        qa: average(items, (item) => item.overallWeights?.qa),
      },
      transferRatePercent: average(items, (item) => item.transferRatePercent),
      admitsCount: average(items, (item) => item.admitsCount),
      firstTimeCaller: average(items, (item) => item.firstTimeCaller),
      transferCount: average(items, (item) => item.transferCount),
      inboundCalls: average(items, (item) => item.inboundCalls),
      inboundMinutesSeconds: average(items, (item) => item.inboundMinutesSeconds),
      holdTimeSeconds: average(items, (item) => item.holdTimeSeconds),
      ahtSeconds: average(items, (item) => item.ahtSeconds),
      attendancePercentValue: average(items, (item) => item.attendancePercentValue),
      qaPercentValue: average(items, (item) => item.qaPercentValue),
      agentCount: items.length,
    }))
    .sort((left, right) => (left.weekDate?.getTime() ?? 0) - (right.weekDate?.getTime() ?? 0));

  return {
    records,
    weeklyAverages,
    monthOptions: [...new Set(records.map((record) => record.monthLabel))],
    weekOptions: [...new Set(records.map((record) => record.weekEnding))],
    agentOptions: [...new Set(records.map((record) => record.agentName))].sort((left, right) =>
      left.localeCompare(right)
    ),
  };
}

export function buildRealtimeDataset(rawDatasets) {
  const primaryKeyLookup = buildPrimaryKeyLookup(rawDatasets.primaryKey || []);
  const aggregatedRealtimeRows = aggregateRealtimeRows(rawDatasets.realtime || [], primaryKeyLookup);
  const dailyRecords = aggregateRealtimeDailyRows(rawDatasets.realtime || [], primaryKeyLookup)
    .sort((left, right) => {
      const leftTime = left.dateAt?.getTime() ?? 0;
      const rightTime = right.dateAt?.getTime() ?? 0;
      return leftTime - rightTime || left.agentName.localeCompare(right.agentName);
    });
  const records = aggregatedRealtimeRows
    .map((row) => {
      const dateValue = String(row.dateRange || row.date || "").trim();
      if (!dateValue) return null;

      const person = resolvePerson(primaryKeyLookup, row.email, row.agentName);
      const weekDate = parseDateRange(dateValue);
      if (!person.identity || !weekDate) return null;

      const firstTimeCaller = safeNumber(row.firstTimeCaller);
      const transferCount = safeNumber(row.transferCount);
      const inboundCalls = safeNumber(row.inboundCalls);
      const transferRate = calculateTransferRate(transferCount, firstTimeCaller);
      const transferScore = calculateTransferScore(transferRate);
      const ahtSeconds = calculateAHT(row.inboundMinutes, row.holdTime, inboundCalls);
      const ahtScore = calculateAHTScore(ahtSeconds);
      const performanceScore = calculatePerformanceScore(transferScore, null, ahtScore);
      const overallComposition = getOverallComposition(performanceScore, null, null);
      const overallScore = calculateOverallScore(performanceScore, null, null);
      const lastUpdatedAt = row.lastUpdatedAt instanceof Date ? row.lastUpdatedAt : parseTimestamp(row.lastUpdated);

      return {
        key: makeKey(person.identity, dateValue),
        identity: person.identity,
        email: person.email || normalizeEmail(row.email),
        agentName: person.displayName || normalizeDisplayName(row.agentName),
        weekEnding: dateValue,
        weekDate,
        monthLabel: titleCaseMonth(weekDate),
        dateRange: dateValue,
        firstTimeCaller,
        transferCount,
        inboundCalls,
        inboundMinutes: row.inboundMinutes || "",
        holdTime: row.holdTime || "",
        inboundMinutesSeconds: secondsFromDuration(row.inboundMinutes),
        holdTimeSeconds: secondsFromDuration(row.holdTime),
        transferRate,
        transferRatePercent: transferRate === null ? null : transferRate * 100,
        transferRateDisplay: formatPercent(transferRate === null ? null : transferRate * 100),
        transferScore,
        admitsCount: null,
        admitsScore: null,
        ahtSeconds,
        ahtDisplay: formatAHT(ahtSeconds),
        ahtScore,
        attendancePercentValue: null,
        attendancePercentDisplay: "N/A",
        attendanceScore: null,
        qaPercentValue: null,
        qaPercentDisplay: "N/A",
        qaScore: null,
        performanceScore,
        overallScore,
        overallIncludesQa: false,
        overallWeights: overallComposition.weights,
        lastUpdated: row.lastUpdated || "",
        lastUpdatedAt,
        lastUpdatedDisplay: formatLastUpdatedDisplay(row.lastUpdated, lastUpdatedAt),
      };
    })
    .filter(Boolean)
    .sort((left, right) => {
      const leftTime = left.weekDate?.getTime() ?? 0;
      const rightTime = right.weekDate?.getTime() ?? 0;
      return leftTime - rightTime || left.agentName.localeCompare(right.agentName);
    });

  const weekGroups = groupByWeek(records);
  const weeklyAverages = [...weekGroups.entries()]
    .map(([weekEnding, items]) => {
      const latestLastUpdated = [...items]
        .map((item) => item.lastUpdatedAt)
        .filter((value) => value instanceof Date && !Number.isNaN(value.getTime()))
        .sort((left, right) => left.getTime() - right.getTime())
        .at(-1) || null;
      const latestLastUpdatedRecord = [...items]
        .filter((item) => item.lastUpdatedAt instanceof Date && !Number.isNaN(item.lastUpdatedAt.getTime()))
        .sort((left, right) => left.lastUpdatedAt.getTime() - right.lastUpdatedAt.getTime())
        .at(-1) || null;

      return {
        weekEnding,
        weekDate: items[0]?.weekDate || parseDateRange(weekEnding),
        monthLabel: items[0]?.monthLabel || "Unknown",
        transferScore: average(items, (item) => item.transferScore),
        admitsScore: null,
        ahtScore: average(items, (item) => item.ahtScore),
        attendanceScore: null,
        qaScore: null,
        performanceScore: average(items, (item) => item.performanceScore),
        overallScore: average(items, (item) => item.overallScore),
        overallIncludesQa: false,
        overallWeights: {
          performance: average(items, (item) => item.overallWeights?.performance),
          attendance: 0,
          qa: 0,
        },
        transferRatePercent: average(items, (item) => item.transferRatePercent),
        admitsCount: null,
        firstTimeCaller: average(items, (item) => item.firstTimeCaller),
        transferCount: average(items, (item) => item.transferCount),
        inboundCalls: average(items, (item) => item.inboundCalls),
        inboundMinutesSeconds: average(items, (item) => item.inboundMinutesSeconds),
        holdTimeSeconds: average(items, (item) => item.holdTimeSeconds),
        ahtSeconds: average(items, (item) => item.ahtSeconds),
        attendancePercentValue: null,
        qaPercentValue: null,
        agentCount: items.length,
        lastUpdatedAt: latestLastUpdated,
        lastUpdatedDisplay: formatLastUpdatedDisplay(latestLastUpdatedRecord?.lastUpdated, latestLastUpdated),
      };
    })
    .sort((left, right) => (left.weekDate?.getTime() ?? 0) - (right.weekDate?.getTime() ?? 0));

  const latestLastUpdated = [...records]
    .map((record) => record.lastUpdatedAt)
    .filter((value) => value instanceof Date && !Number.isNaN(value.getTime()))
    .sort((left, right) => left.getTime() - right.getTime())
    .at(-1) || null;
  const latestLastUpdatedRecord = [...records]
    .filter((record) => record.lastUpdatedAt instanceof Date && !Number.isNaN(record.lastUpdatedAt.getTime()))
    .sort((left, right) => left.lastUpdatedAt.getTime() - right.lastUpdatedAt.getTime())
    .at(-1) || null;

  return {
    records,
    dailyRecords,
    weeklyAverages,
    monthOptions: [...new Set(records.map((record) => record.monthLabel))],
    weekOptions: [...new Set(records.map((record) => record.weekEnding))],
    agentOptions: [...new Set(records.map((record) => record.agentName))].sort((left, right) =>
      left.localeCompare(right)
    ),
    lastUpdatedAt: latestLastUpdated,
    lastUpdatedDisplay: formatLastUpdatedDisplay(latestLastUpdatedRecord?.lastUpdated, latestLastUpdated),
  };
}
