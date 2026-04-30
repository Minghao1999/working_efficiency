export function fileIdentity(file) {
  return `${file.name}|${file.size ?? 0}|${file.lastModified ?? 0}`;
}

export function mergeFiles(current, nextFiles) {
  const map = new Map(current.map((file) => [fileIdentity(file), file]));
  for (const file of nextFiles) {
    map.set(fileIdentity(file), file);
  }
  return [...map.values()];
}

export function isLikelyVolumeFile(file) {
  return /件量|volume|sales|销售单|order/i.test(file.name || "");
}

export function mergeEfficiencyVolumeDelta(current, delta) {
  if (!current) return current;
  const days = { ...current.days };
  for (const [date, patch] of Object.entries(delta.days || {})) {
    if (!days[date]) continue;
    days[date] = {
      ...days[date],
      kpi: {
        ...days[date].kpi,
        volume: patch.kpi?.volume ?? null
      }
    };
  }
  return {
    ...current,
    completeness: { ...current.completeness, ...delta.completeness },
    days
  };
}
