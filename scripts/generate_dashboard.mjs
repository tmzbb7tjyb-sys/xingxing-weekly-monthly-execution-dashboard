#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";

const METRICS = [
  ["samples", "寄样数"],
  ["scripts", "脚本数"],
  ["publishes", "发布数"],
  ["amount", "落地金额"],
];

const COLORS = {
  navy: "#1F4E78",
  blue: "#4472C4",
  paleBlue: "#DDEBF7",
  input: "#FFF2CC",
  paleGreen: "#E2F0D9",
  paleRed: "#FCE4D6",
  paleGray: "#F2F2F2",
  grid: "#D9E2F3",
  text: "#1F2937",
  red: "#C00000",
  white: "#FFFFFF",
};

const CURRENCY_FORMAT = '"¥"#,##0';

function usage() {
  return [
    "Usage:",
    "  generate_dashboard.mjs --config /absolute/path/config.json [--output /absolute/path/output.xlsx] [--force]",
  ].join("\n");
}

function parseArgs(argv) {
  const result = { force: false };
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "--config" || arg === "--output") {
      const value = argv[i + 1];
      if (!value) throw new Error(`${arg} requires a value.`);
      result[arg.slice(2)] = value;
      i += 1;
    } else if (arg === "--force") {
      result.force = true;
    } else if (arg === "--help" || arg === "-h") {
      console.log(usage());
      process.exit(0);
    } else {
      throw new Error(`Unknown argument: ${arg}`);
    }
  }
  if (!result.config) throw new Error("--config is required.");
  return result;
}

function isBlank(value) {
  return value === null || value === undefined || value === "";
}

function normalizeNumber(value, label) {
  if (isBlank(value)) return null;
  if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
    throw new Error(`${label} must be a non-negative number or blank.`);
  }
  return value;
}

function normalizeMetricObject(value, label) {
  const source = value && typeof value === "object" && !Array.isArray(value) ? value : {};
  return Object.fromEntries(
    METRICS.map(([key]) => [key, normalizeNumber(source[key], `${label}.${key}`)]),
  );
}

function normalizeText(value) {
  return isBlank(value) ? "" : String(value);
}

function normalizeTask(task, taskIndex, weekCount) {
  if (!task || typeof task !== "object" || Array.isArray(task)) {
    throw new Error(`tasks[${taskIndex}] must be an object.`);
  }
  const name = normalizeText(task.name).trim();
  if (!name) throw new Error(`tasks[${taskIndex}].name is required.`);

  const weeklyPlans = Array.isArray(task.weekly_plans) ? task.weekly_plans : [];
  const actuals = Array.isArray(task.actuals) ? task.actuals : [];
  const feedback = Array.isArray(task.feedback) ? task.feedback : [];
  for (const [field, entries] of [["weekly_plans", weeklyPlans], ["actuals", actuals], ["feedback", feedback]]) {
    if (entries.length > weekCount) {
      throw new Error(`tasks[${taskIndex}].${field} has ${entries.length} items, but the month has ${weekCount} weeks.`);
    }
  }

  return {
    name,
    brand: normalizeText(task.brand),
    product: normalizeText(task.product),
    pm: normalizeText(task.pm),
    taskId: normalizeText(task.task_id),
    monthlyTargets: normalizeMetricObject(task.monthly_targets, `tasks[${taskIndex}].monthly_targets`),
    weeklyPlans: Array.from({ length: weekCount }, (_, i) =>
      normalizeMetricObject(weeklyPlans[i], `tasks[${taskIndex}].weekly_plans[${i}]`),
    ),
    actuals: Array.from({ length: weekCount }, (_, i) => {
      const week = actuals[i] && typeof actuals[i] === "object" ? actuals[i] : {};
      return {
        mid: normalizeMetricObject(week.mid, `tasks[${taskIndex}].actuals[${i}].mid`),
        tail: normalizeMetricObject(week.tail, `tasks[${taskIndex}].actuals[${i}].tail`),
        full: normalizeMetricObject(week.full, `tasks[${taskIndex}].actuals[${i}].full`),
      };
    }),
    feedback: Array.from({ length: weekCount }, (_, i) => {
      const week = feedback[i] && typeof feedback[i] === "object" ? feedback[i] : {};
      return {
        mid: normalizeText(week.mid),
        tail: normalizeText(week.tail),
        full: normalizeText(week.full),
      };
    }),
  };
}

function normalizeCheckpoint(value, weekCount) {
  const source = value && typeof value === "object" && !Array.isArray(value) ? value : {};
  const mode = source.mode === "manual" ? "manual" : "auto";
  const week = Number(source.week ?? 1);
  if (!Number.isInteger(week) || week < 1 || week > weekCount) {
    throw new Error(`checkpoint.week must be an integer between 1 and ${weekCount}.`);
  }
  const stageMap = {
    mid: "周中",
    tail: "周尾",
    full: "整周",
    周中: "周中",
    周尾: "周尾",
    整周: "整周",
  };
  const stage = stageMap[source.stage ?? "mid"];
  if (!stage) throw new Error("checkpoint.stage must be mid, tail, full, 周中, 周尾, or 整周.");
  return { mode, week, stage };
}

function monthParts(month) {
  if (!/^\d{6}$/.test(month)) throw new Error("month must use YYYYMM format.");
  const year = Number(month.slice(0, 4));
  const monthNumber = Number(month.slice(4, 6));
  if (year < 2000 || year > 2099 || monthNumber < 1 || monthNumber > 12) {
    throw new Error("month must be a valid month between 200001 and 209912.");
  }
  return { year, monthNumber };
}

function buildWeeks(year, monthNumber) {
  const daysInMonth = new Date(Date.UTC(year, monthNumber, 0)).getUTCDate();
  const weeks = [];
  let day = 1;
  while (day <= daysInMonth) {
    const date = new Date(Date.UTC(year, monthNumber - 1, day));
    const weekday = date.getUTCDay();
    const daysToSunday = (7 - weekday) % 7;
    const endDay = Math.min(daysInMonth, day + daysToSunday);
    weeks.push({ startDay: day, endDay });
    day = endDay + 1;
  }
  return { weeks, daysInMonth };
}

function columnName(index) {
  let n = index;
  let result = "";
  while (n > 0) {
    n -= 1;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}

function weekdayName(year, monthNumber, day) {
  return ["周日", "周一", "周二", "周三", "周四", "周五", "周六"][
    new Date(Date.UTC(year, monthNumber - 1, day)).getUTCDay()
  ];
}

function weekTitle(index, week, year, monthNumber) {
  const ordinals = ["第一周", "第二周", "第三周", "第四周", "第五周", "第六周"];
  const mm = String(monthNumber).padStart(2, "0");
  const start = `${mm}${String(week.startDay).padStart(2, "0")}`;
  const end = `${mm}${String(week.endDay).padStart(2, "0")}`;
  return `${ordinals[index]}（${start}-${end}）${weekdayName(year, monthNumber, week.startDay)}-${weekdayName(year, monthNumber, week.endDay)}`;
}

function safeSheetName(base, usedNames) {
  const cleaned = base.replace(/[\\/?*\[\]:]/g, "-").trim() || "任务";
  let candidate = cleaned.slice(0, 31);
  let suffix = 2;
  while (usedNames.has(candidate)) {
    const tag = `-${suffix}`;
    candidate = `${cleaned.slice(0, 31 - tag.length)}${tag}`;
    suffix += 1;
  }
  usedNames.add(candidate);
  return candidate;
}

function indirectRef(sheetName, cellAddress) {
  const escapedName = sheetName.replace(/'/g, "''").replace(/"/g, '""');
  return `INDIRECT("'${escapedName}'!${cellAddress}")`;
}

function linkedTextOrBlank(source) {
  return `=IF(ISBLANK(${source}),"",${source}&"")`;
}

function linkedNumberOrBlank(source) {
  return `=IF(COUNT(${source})=0,"",SUM(${source}))`;
}

function linkedNumberOrTextOrBlank(source) {
  return `=IF(COUNT(${source})>0,SUM(${source}),IF(ISBLANK(${source}),"",${source}&""))`;
}

function latestNumericRefsOrBlank(refs) {
  const [mid, tail, full] = refs;
  return `=IF(COUNT(${full})>0,SUM(${full}),IF(COUNT(${tail})>0,SUM(${tail}),IF(COUNT(${mid})>0,SUM(${mid}),"")))`;
}

function latestValueExpression(mid, tail, full) {
  return `IF(COUNT(${full})>0,SUM(${full}),IF(COUNT(${tail})>0,SUM(${tail}),IF(COUNT(${mid})>0,SUM(${mid}),0)))`;
}

function latestCountExpression(mid, tail, full) {
  return `IF(COUNT(${full})>0,1,IF(COUNT(${tail})>0,1,IF(COUNT(${mid})>0,1,0)))`;
}

function checkpointValueExpression(mid, tail, full, stageRef) {
  return `IF(${stageRef}="周中",IF(COUNT(${mid})>0,SUM(${mid}),0),IF(${stageRef}="周尾",IF(COUNT(${tail})>0,SUM(${tail}),0),IF(${stageRef}="整周",IF(COUNT(${full})>0,SUM(${full}),0),0)))`;
}

function checkpointCountExpression(mid, tail, full, stageRef) {
  return `IF(${stageRef}="周中",IF(COUNT(${mid})>0,1,0),IF(${stageRef}="周尾",IF(COUNT(${tail})>0,1,0),IF(${stageRef}="整周",IF(COUNT(${full})>0,1,0),0)))`;
}

function checkpointActualFormula(refs, weekNumber, controls) {
  const [mid, tail, full] = refs;
  const checkpointCount = checkpointCountExpression(mid, tail, full, controls.stageRef);
  return `=IF(${controls.weekRef}=${weekNumber},IF(${checkpointCount}=0,"",${checkpointValueExpression(mid, tail, full, controls.stageRef)}),IF(${latestCountExpression(mid, tail, full)}=0,"",${latestValueExpression(mid, tail, full)}))`;
}

function textMatrix(value) {
  return [[value]];
}

function applyBaseStyle(sheet, lastColumn, lastRow) {
  const full = sheet.getRange(`A1:${lastColumn}${lastRow}`);
  full.format.font = { name: "Arial", size: 10, color: COLORS.text };
  full.format.verticalAlignment = "center";
  sheet.showGridLines = false;
}

function addStatusFormatting(range) {
  range.conditionalFormats.add("containsText", {
    text: "滞后",
    format: { fill: COLORS.paleRed, font: { color: COLORS.red, bold: true } },
  });
  range.conditionalFormats.add("containsText", {
    text: "正常",
    format: { fill: COLORS.paleGreen, font: { color: "#375623", bold: true } },
  });
  range.conditionalFormats.add("containsText", {
    text: "超预期",
    format: { fill: COLORS.paleBlue, font: { color: COLORS.navy, bold: true } },
  });
  range.conditionalFormats.add("containsText", {
    text: "待填写",
    format: { fill: COLORS.input, font: { color: "#7F6000", bold: true } },
  });
  range.conditionalFormats.add("containsText", {
    text: "计划待补",
    format: { fill: COLORS.input, font: { color: "#7F6000", bold: true } },
  });
  range.conditionalFormats.add("containsText", {
    text: "无计划",
    format: { fill: COLORS.paleGray, font: { color: "#595959", bold: true } },
  });
  range.conditionalFormats.add("containsText", {
    text: "计划外完成",
    format: { fill: COLORS.paleBlue, font: { color: COLORS.navy, bold: true } },
  });
}

function stageComponents(weeks, planRefs, actualRefs, controls) {
  const planWeights = weeks.map((_, i) => `IF(${controls.weekRef}>${i + 1},1,IF(${controls.weekRef}=${i + 1},N(${controls.ratioRef}),0))`);
  return {
    expectedCount: controls.weekRef,
    planCount: planRefs.map((ref, i) => `IF(${controls.weekRef}>=${i + 1},COUNT(${ref}),0)`).join("+"),
    planValue: planRefs.map((ref, i) => `SUM(${ref})*(${planWeights[i]})`).join("+"),
    actualCount: actualRefs.map((refs, i) => {
      const [mid, tail, full] = refs;
      return `IF(${controls.weekRef}>${i + 1},${latestCountExpression(mid, tail, full)},IF(${controls.weekRef}=${i + 1},${checkpointCountExpression(mid, tail, full, controls.stageRef)},0))`;
    }).join("+"),
    actualValue: actualRefs.map((refs, i) => {
      const [mid, tail, full] = refs;
      return `IF(${controls.weekRef}>${i + 1},${latestValueExpression(mid, tail, full)},IF(${controls.weekRef}=${i + 1},${checkpointValueExpression(mid, tail, full, controls.stageRef)},0))`;
    }).join("+"),
  };
}

function stagePlanReadyExpression(parts) {
  return `((${parts.planCount})=(${parts.expectedCount}))`;
}

function stageActualReadyExpression(parts) {
  return `((${parts.actualCount})=(${parts.expectedCount}))`;
}

function rateStatusExpression(rate, suffix = "") {
  const suffixText = suffix ? `&"${suffix}"` : "";
  return `IF(${rate}<0.95,"滞后 "&TEXT(${rate},"0%")${suffixText},IF(${rate}<=1.05,"正常 "&TEXT(${rate},"0%")${suffixText},"超预期 "&TEXT(${rate},"0%")${suffixText}))`;
}

function singleStageStatusFormula(parts) {
  const rate = `((${parts.actualValue})/(${parts.planValue}))`;
  return `=IF(${parts.stageRef}="待填写","待填写",IF(NOT(${stagePlanReadyExpression(parts)}),"计划待补",IF(NOT(${stageActualReadyExpression(parts)}),"待填写",IF((${parts.planValue})=0,IF((${parts.actualValue})=0,"无计划","计划外完成"),${rateStatusExpression(rate)}))))`;
}

function combinedStageStatusFormula(first, second) {
  const firstPlanMissing = `NOT(${stagePlanReadyExpression(first)})`;
  const secondPlanMissing = `NOT(${stagePlanReadyExpression(second)})`;
  const firstActualMissing = `NOT(${stageActualReadyExpression(first)})`;
  const secondActualMissing = `NOT(${stageActualReadyExpression(second)})`;
  const firstRate = `((${first.actualValue})/(${first.planValue}))`;
  const secondRate = `((${second.actualValue})/(${second.planValue}))`;
  const bothPlansZero = `AND((${first.planValue})=0,(${second.planValue})=0)`;
  const weakRate = `MIN(${firstRate},${secondRate})`;
  const weakSuffix = `IF(ABS(${firstRate}-${secondRate})<0.0000001,"",IF(${firstRate}<${secondRate},"（发布）","（金额）"))`;
  const bothRateStatus = `IF(${weakRate}<0.95,"滞后 "&TEXT(${weakRate},"0%")&${weakSuffix},IF(${weakRate}<=1.05,"正常 "&TEXT(${weakRate},"0%")&${weakSuffix},"超预期 "&TEXT(${weakRate},"0%")&${weakSuffix}))`;
  const firstOnlyStatus = rateStatusExpression(firstRate, "（发布）");
  const secondOnlyStatus = rateStatusExpression(secondRate, "（金额）");
  return `=IF(${first.stageRef}="待填写","待填写",IF(OR(${firstPlanMissing},${secondPlanMissing}),IF(AND(${firstPlanMissing},${secondPlanMissing}),"计划待补",IF(${firstPlanMissing},"计划待补（发布）","计划待补（金额）")),IF(OR(${firstActualMissing},${secondActualMissing}),IF(AND(${firstActualMissing},${secondActualMissing}),"待填写",IF(${firstActualMissing},"待填写（发布）","待填写（金额）")),IF(${bothPlansZero},IF(AND((${first.actualValue})=0,(${second.actualValue})=0),"无计划","计划外完成"),IF((${first.planValue})=0,${secondOnlyStatus},IF((${second.planValue})=0,${firstOnlyStatus},${bothRateStatus}))))))`;
}

function buildTaskSheet(sheet, task, weeks, year, monthNumber, daysInMonth, controls) {
  const lastColIndex = 8 + weeks.length * 7;
  const lastCol = columnName(lastColIndex);
  applyBaseStyle(sheet, lastCol, 9);
  sheet.tabColor = COLORS.blue;

  sheet.mergeCells("B1:C1");
  sheet.mergeCells("D1:H1");
  sheet.getRange("A1").values = textMatrix("PM");
  sheet.getRange("B1").values = textMatrix(task.pm);
  sheet.getRange("D1").values = textMatrix("月度进度");

  const fixedHeaders = ["品牌", "产品", "星图任务ID", "数据维度", "月度规划", "月度完成", "月度完成率", "进度判断"];
  sheet.getRange("A2:H2").values = [fixedHeaders];
  sheet.mergeCells("A3:A6");
  sheet.mergeCells("B3:B6");
  sheet.mergeCells("C3:C6");
  sheet.getRange("C3:C6").format.numberFormat = "@";
  sheet.getRange("A3").values = textMatrix(task.brand);
  sheet.getRange("B3").values = textMatrix(task.product);
  sheet.getRange("C3").values = textMatrix(task.taskId ? `'${task.taskId}` : "");
  sheet.getRange("D3:D6").values = METRICS.map(([, label]) => [label]);
  sheet.getRange("E3:E6").values = METRICS.map(([key]) => [task.monthlyTargets[key]]);

  for (let metricIndex = 0; metricIndex < METRICS.length; metricIndex += 1) {
    const row = 3 + metricIndex;
    const helperCols = Array.from({ length: 4 }, (_, i) => columnName(metricIndex * 4 + 1 + i));
    const actualRefs = [];
    const weeklyLatest = [];
    const weeklyActualRefs = [];
    const planRefs = [];
    for (let w = 0; w < weeks.length; w += 1) {
      const base = 9 + w * 7;
      const plan = `${columnName(base)}${row}`;
      const mid = `${columnName(base + 1)}${row}`;
      const tail = `${columnName(base + 3)}${row}`;
      const full = `${columnName(base + 5)}${row}`;
      planRefs.push(plan);
      actualRefs.push(mid, tail, full);
      weeklyLatest.push(latestValueExpression(mid, tail, full));
      weeklyActualRefs.push([mid, tail, full]);
    }
    sheet.getRange(`F${row}`).formulas = [[`=IF(COUNT(${actualRefs.join(",")})=0,"",${weeklyLatest.join("+")})`]];
    sheet.getRange(`G${row}`).formulas = [[`=IF(COUNT(E${row})=0,"",IF(E${row}=0,"n.a.",IF(COUNT(F${row})=0,"",F${row}/E${row})))`]];
    const stage = stageComponents(weeks, planRefs, weeklyActualRefs, controls);
    sheet.getRange(`${helperCols[0]}8:${helperCols[3]}8`).formulas = [[
      `=${stage.planCount}`,
      `=${stage.planValue}`,
      `=${stage.actualCount}`,
      `=${stage.actualValue}`,
    ]];
    const helperStage = {
      stageRef: controls.stageRef,
      expectedCount: controls.weekRef,
      planCount: `${helperCols[0]}8`,
      planValue: `${helperCols[1]}8`,
      actualCount: `${helperCols[2]}8`,
      actualValue: `${helperCols[3]}8`,
    };
    sheet.getRange(`H${row}`).formulas = [[singleStageStatusFormula(helperStage)]];
  }

  const weeklyHeaders = ["单周规划", "周中完成", "周中进度", "周尾完成", "周尾进度", "整周实际", "整周进度"];
  const feedbackDefaults = {
    mid: "周中反馈（周四早）\n填写：进度、差距、动作、负责人、完成时间、验证数字",
    tail: "周尾反馈（周五/六晚）\n填写：进度、差距、动作、负责人、完成时间、验证数字",
    full: "整周反馈（次周一早）\n填写：进度、差距、动作、负责人、完成时间、验证数字",
  };

  for (let w = 0; w < weeks.length; w += 1) {
    const base = 9 + w * 7;
    const startCol = columnName(base);
    const endCol = columnName(base + 6);
    sheet.mergeCells(`${startCol}1:${endCol}1`);
    sheet.getRange(`${startCol}1`).values = textMatrix(weekTitle(w, weeks[w], year, monthNumber));
    sheet.getRange(`${startCol}2:${endCol}2`).values = [weeklyHeaders];

    for (let metricIndex = 0; metricIndex < METRICS.length; metricIndex += 1) {
      const [key] = METRICS[metricIndex];
      const row = 3 + metricIndex;
      const planCol = columnName(base);
      const midCol = columnName(base + 1);
      const midRateCol = columnName(base + 2);
      const tailCol = columnName(base + 3);
      const tailRateCol = columnName(base + 4);
      const fullCol = columnName(base + 5);
      const fullRateCol = columnName(base + 6);
      sheet.getRange(`${planCol}${row}`).values = [[task.weeklyPlans[w][key]]];
      sheet.getRange(`${midCol}${row}`).values = [[task.actuals[w].mid[key]]];
      sheet.getRange(`${tailCol}${row}`).values = [[task.actuals[w].tail[key]]];
      sheet.getRange(`${fullCol}${row}`).values = [[task.actuals[w].full[key]]];
      for (const [actualCol, rateCol] of [[midCol, midRateCol], [tailCol, tailRateCol], [fullCol, fullRateCol]]) {
        sheet.getRange(`${rateCol}${row}`).formulas = [[`=IF(COUNT(${planCol}${row})=0,"",IF(${planCol}${row}=0,"n.a.",IF(COUNT(${actualCol}${row})=0,"",${actualCol}${row}/${planCol}${row})))`]];
      }
    }

    const midStart = columnName(base + 1);
    const midEnd = columnName(base + 2);
    const tailStart = columnName(base + 3);
    const tailEnd = columnName(base + 4);
    const fullStart = columnName(base + 5);
    const fullEnd = columnName(base + 6);
    sheet.mergeCells(`${midStart}7:${midEnd}7`);
    sheet.mergeCells(`${tailStart}7:${tailEnd}7`);
    sheet.mergeCells(`${fullStart}7:${fullEnd}7`);
    sheet.getRange(`${midStart}7`).values = [[task.feedback[w].mid || feedbackDefaults.mid]];
    sheet.getRange(`${tailStart}7`).values = [[task.feedback[w].tail || feedbackDefaults.tail]];
    sheet.getRange(`${fullStart}7`).values = [[task.feedback[w].full || feedbackDefaults.full]];

    sheet.getRange(`${startCol}3:${startCol}6`).format.fill = COLORS.input;
    for (const actualOffset of [1, 3, 5]) {
      const actualCol = columnName(base + actualOffset);
      sheet.getRange(`${actualCol}3:${actualCol}6`).format.fill = COLORS.paleBlue;
    }
    sheet.getRange(`${midStart}7:${fullEnd}7`).format.fill = "#EAF3F8";
    sheet.getRange(`${midStart}7:${fullEnd}7`).format.wrapText = true;
    sheet.getRange(`${startCol}1:${endCol}2`).format.fill = COLORS.blue;
    sheet.getRange(`${startCol}1:${endCol}2`).format.font = { name: "Arial", size: 10, bold: true, color: COLORS.white };
    sheet.getRange(`${startCol}1:${endCol}2`).format.horizontalAlignment = "center";

    const thresholdRules = [
      [columnName(base + 2), 0.5],
      [columnName(base + 4), 0.85],
      [columnName(base + 6), 1],
    ];
    for (const [rateCol, threshold] of thresholdRules) {
      const rateRange = sheet.getRange(`${rateCol}3:${rateCol}6`);
      rateRange.conditionalFormats.add("dataBar", {
        color: COLORS.blue,
        thresholds: [{ type: "num", value: 0 }, { type: "num", value: 1 }],
        gradient: true,
      });
      rateRange.conditionalFormats.addCustom(
        `=AND(COUNT(${rateCol}3)=1,${rateCol}3<${threshold})`,
        { font: { color: COLORS.red, bold: true } },
      );
    }
  }

  const planCells = weeks.map((_, w) => columnName(9 + w * 7));
  const checks = METRICS.map((_, i) => {
    const row = 3 + i;
    return `E${row}=SUM(${planCells.map((col) => `${col}${row}`).join(",")})`;
  });
  sheet.mergeCells("A7:D7");
  sheet.mergeCells("E7:G7");
  sheet.getRange("A7").values = textMatrix("月目标校验");
  sheet.getRange("E7").formulas = [[`=IF(COUNT(E3:E6)<4,"待填写",IF(AND(${checks.join(",")}),"4项一致","请检查月度规划与周计划"))`]];
  sheet.getRange("H7").values = textMatrix("情况说明");

  sheet.getRange("A1:H2").format.fill = COLORS.navy;
  sheet.getRange("A1:H2").format.font = { name: "Arial", size: 10, bold: true, color: COLORS.white };
  sheet.getRange(`A1:${lastCol}2`).format.horizontalAlignment = "center";
  sheet.getRange("A3:C6").format.fill = COLORS.input;
  sheet.getRange("E3:E6").format.fill = COLORS.input;
  sheet.getRange("G3:G6").format.numberFormat = "0.0%";
  sheet.getRange("G3:G6").conditionalFormats.add("dataBar", {
    color: COLORS.blue,
    thresholds: [{ type: "num", value: 0 }, { type: "num", value: 1 }],
    gradient: true,
  });
  addStatusFormatting(sheet.getRange("H3:H6"));
  sheet.getRange("E7:G7").conditionalFormats.add("containsText", {
    text: "请检查",
    format: { fill: COLORS.paleRed, font: { color: COLORS.red, bold: true } },
  });
  sheet.getRange("E7:G7").conditionalFormats.add("containsText", {
    text: "4项一致",
    format: { fill: COLORS.paleGreen, font: { color: "#375623", bold: true } },
  });
  sheet.getRange("E7:G7").conditionalFormats.add("containsText", {
    text: "待填写",
    format: { fill: COLORS.input, font: { color: "#7F6000", bold: true } },
  });
  sheet.getRange(`A7:${lastCol}7`).format.borders = { preset: "all", style: "thin", color: COLORS.grid };
  sheet.getRange(`A1:${lastCol}7`).format.borders = { preset: "all", style: "thin", color: COLORS.grid };
  sheet.getRange(`D3:${lastCol}6`).format.horizontalAlignment = "center";
  sheet.getRange("A3:C6").format.horizontalAlignment = "left";
  sheet.getRange("E3:F6").format.numberFormat = "#,##0";
  sheet.getRange("E6:F6").format.numberFormat = CURRENCY_FORMAT;
  sheet.getRange(`I3:${lastCol}6`).format.numberFormat = "#,##0";
  for (let w = 0; w < weeks.length; w += 1) {
    const base = 9 + w * 7;
    for (const offset of [0, 1, 3, 5]) {
      const col = columnName(base + offset);
      sheet.getRange(`${col}6`).format.numberFormat = CURRENCY_FORMAT;
    }
    for (const offset of [2, 4, 6]) {
      const col = columnName(base + offset);
      sheet.getRange(`${col}3:${col}6`).format.numberFormat = "0.0%";
    }
  }

  sheet.mergeCells("B9:D9");
  sheet.mergeCells("E9:G9");
  sheet.mergeCells(`H9:${lastCol}9`);
  sheet.getRange("A9").values = textMatrix("填写说明");
  sheet.getRange("B9").values = textMatrix("浅黄色：目标和计划");
  sheet.getRange("E9").values = textMatrix("浅蓝色：累计实际");
  sheet.getRange("H9").values = textMatrix("周中、周尾、整周均填写本周累计值；汇总页按日期和已填写阶段自动识别检查点。已结束周计100%，当前周按周中50%/周尾85%/整周100%计入，95%-105%为正常。周中低于50%、周尾低于85%、整周低于100%时红字提醒。");
  sheet.getRange(`A9:${lastCol}9`).format.fill = COLORS.paleGray;
  sheet.getRange(`A9:${lastCol}9`).format.wrapText = true;
  sheet.getRange(`A9:${lastCol}9`).format.borders = { preset: "outside", style: "thin", color: COLORS.grid };

  const widths = { A: 12, B: 18, C: 23, D: 12, E: 13, F: 13, G: 13, H: 14 };
  for (const [col, width] of Object.entries(widths)) sheet.getRange(`${col}1:${col}9`).format.columnWidth = width;
  for (let w = 0; w < weeks.length; w += 1) {
    const base = 9 + w * 7;
    const weekWidths = [12, 12, 11, 12, 11, 12, 11];
    for (let offset = 0; offset < 7; offset += 1) {
      const col = columnName(base + offset);
      sheet.getRange(`${col}1:${col}9`).format.columnWidth = weekWidths[offset];
    }
  }
  sheet.getRange(`A1:${lastCol}1`).format.rowHeight = 28;
  sheet.getRange(`A2:${lastCol}2`).format.rowHeight = 38;
  sheet.getRange(`A3:${lastCol}6`).format.rowHeight = 28;
  sheet.getRange(`A7:${lastCol}7`).format.rowHeight = 88;
  sheet.getRange(`A8:${lastCol}8`).format.rowHeight = 8;
  sheet.getRange("A8:P8").format.font = { name: "Arial", size: 1, color: COLORS.white };
  sheet.getRange("A8:P8").format.fill = COLORS.white;
  sheet.getRange(`A9:${lastCol}9`).format.rowHeight = 42;
  sheet.freezePanes.freezeRows(2);
  sheet.freezePanes.freezeColumns(8);
}

function buildSummarySheet(sheet, taskRecords, weeks, year, monthNumber, daysInMonth, checkpoint, controls) {
  const lastColIndex = 10 + weeks.length * 6;
  const lastCol = columnName(lastColIndex);
  const firstTaskRow = 3;
  const lastTaskRow = firstTaskRow + taskRecords.length - 1;
  const totalRow = lastTaskRow + 1;
  const controlRow = controls.controlRow;
  const autoHelperRow = controlRow + 3;
  const noteRow = controlRow + 4;
  const helperRows = [totalRow + 1, controlRow + 1, controlRow + 2];
  const helperColsForRecord = (recordIndex) => {
    const start = recordIndex * 8;
    const helperRow = helperRows[Math.floor(start / lastColIndex)];
    const startColIndex = (start % lastColIndex) + 1;
    return Array.from({ length: 8 }, (_, i) => `${columnName(startColIndex + i)}${helperRow}`);
  };
  applyBaseStyle(sheet, lastCol, noteRow);
  sheet.tabColor = COLORS.navy;

  for (const col of ["A", "B", "C", "J"]) sheet.mergeCells(`${col}1:${col}2`);
  sheet.mergeCells("D1:F1");
  sheet.mergeCells("G1:I1");
  sheet.getRange("A1").values = textMatrix("品牌");
  sheet.getRange("B1").values = textMatrix("产品");
  sheet.getRange("C1").values = textMatrix("组长/PM");
  sheet.getRange("D1").values = textMatrix("月度发布");
  sheet.getRange("G1").values = textMatrix("月度发布金额");
  sheet.getRange("J1").values = textMatrix("进度状态");
  sheet.getRange("D2:I2").values = [["规划发布数", "完成发布数", "发布完成率", "规划发布金额", "完成发布金额", "金额完成率"]];

  for (let w = 0; w < weeks.length; w += 1) {
    const base = 11 + w * 6;
    const start = columnName(base);
    const end = columnName(base + 5);
    sheet.mergeCells(`${start}1:${end}1`);
    sheet.getRange(`${start}1`).values = [[`W${w + 1}（${String(weeks[w].startDay).padStart(2, "0")}-${String(weeks[w].endDay).padStart(2, "0")}日）`]];
    sheet.getRange(`${start}2:${end}2`).values = [["规划发布数", "完成发布数", "发布完成率", "规划发布金额", "完成发布金额", "金额完成率"]];
    sheet.getRange(`${start}1:${end}2`).format.fill = COLORS.blue;
  }

  for (let i = 0; i < taskRecords.length; i += 1) {
    const row = firstTaskRow + i;
    const helperCells = helperColsForRecord(i);
    const sheetName = taskRecords[i].sheetName;
    const publishPlanRefs = weeks.map((_, w) => `${columnName(11 + w * 6)}${row}`);
    const detailPublishActualRefs = weeks.map((_, w) => {
      const detailBase = 9 + w * 7;
      return [1, 3, 5].map((offset) => indirectRef(sheetName, `${columnName(detailBase + offset)}5`));
    });
    const amountPlanRefs = weeks.map((_, w) => `${columnName(14 + w * 6)}${row}`);
    const detailAmountActualRefs = weeks.map((_, w) => {
      const detailBase = 9 + w * 7;
      return [1, 3, 5].map((offset) => indirectRef(sheetName, `${columnName(detailBase + offset)}6`));
    });
    const publishStage = stageComponents(
      weeks,
      publishPlanRefs,
      detailPublishActualRefs,
      controls,
    );
    const amountStage = stageComponents(
      weeks,
      amountPlanRefs,
      detailAmountActualRefs,
      controls,
    );
    sheet.getRange(`${helperCells[0]}:${helperCells[7]}`).formulas = [[
      `=${publishStage.planCount}`,
      `=${publishStage.planValue}`,
      `=${publishStage.actualCount}`,
      `=${publishStage.actualValue}`,
      `=${amountStage.planCount}`,
      `=${amountStage.planValue}`,
      `=${amountStage.actualCount}`,
      `=${amountStage.actualValue}`,
    ]];
    const publishHelperStage = {
      stageRef: controls.stageRef,
      expectedCount: controls.weekRef,
      planCount: helperCells[0],
      planValue: helperCells[1],
      actualCount: helperCells[2],
      actualValue: helperCells[3],
    };
    const amountHelperStage = {
      stageRef: controls.stageRef,
      expectedCount: controls.weekRef,
      planCount: helperCells[4],
      planValue: helperCells[5],
      actualCount: helperCells[6],
      actualValue: helperCells[7],
    };
    const fixedFormulas = [
      linkedTextOrBlank(indirectRef(sheetName, "A3")),
      linkedTextOrBlank(indirectRef(sheetName, "B3")),
      linkedTextOrBlank(indirectRef(sheetName, "B1")),
      linkedNumberOrBlank(indirectRef(sheetName, "E5")),
      linkedNumberOrBlank(indirectRef(sheetName, "F5")),
      linkedNumberOrTextOrBlank(indirectRef(sheetName, "G5")),
      linkedNumberOrBlank(indirectRef(sheetName, "E6")),
      linkedNumberOrBlank(indirectRef(sheetName, "F6")),
      `=IF(COUNT(G${row})=0,"",IF(G${row}=0,"n.a.",IF(COUNT(H${row})=0,"",H${row}/G${row})))`,
      combinedStageStatusFormula(publishHelperStage, amountHelperStage),
    ];
    sheet.getRange(`A${row}:J${row}`).formulas = [fixedFormulas];

    for (let w = 0; w < weeks.length; w += 1) {
      const summaryBase = 11 + w * 6;
      const detailBase = 9 + w * 7;
      const pubPlan = columnName(detailBase);
      const pubActuals = [columnName(detailBase + 1), columnName(detailBase + 3), columnName(detailBase + 5)];
      const amountPlan = pubPlan;
      const amountActuals = pubActuals;
      const sCols = Array.from({ length: 6 }, (_, offset) => columnName(summaryBase + offset));
      const pubActualRefs = pubActuals.map((col) => indirectRef(sheetName, `${col}5`));
      const amountActualRefs = amountActuals.map((col) => indirectRef(sheetName, `${col}6`));
      sheet.getRange(`${sCols[0]}${row}:${sCols[5]}${row}`).formulas = [[
        linkedNumberOrBlank(indirectRef(sheetName, `${pubPlan}5`)),
        checkpointActualFormula(pubActualRefs, w + 1, controls),
        `=IF(COUNT(${sCols[0]}${row})=0,"",IF(${sCols[0]}${row}=0,"n.a.",IF(COUNT(${sCols[1]}${row})=0,"",${sCols[1]}${row}/${sCols[0]}${row})))`,
        linkedNumberOrBlank(indirectRef(sheetName, `${amountPlan}6`)),
        checkpointActualFormula(amountActualRefs, w + 1, controls),
        `=IF(COUNT(${sCols[3]}${row})=0,"",IF(${sCols[3]}${row}=0,"n.a.",IF(COUNT(${sCols[4]}${row})=0,"",${sCols[4]}${row}/${sCols[3]}${row})))`,
      ]];
    }
  }

  sheet.getRange(`A${totalRow}:C${totalRow}`).values = [["合计", "合计", "合计"]];
  const sumOrBlank = (col) => `=IF(COUNT(${col}${firstTaskRow}:${col}${lastTaskRow})=0,"",SUM(${col}${firstTaskRow}:${col}${lastTaskRow}))`;
  sheet.getRange(`D${totalRow}`).formulas = [[sumOrBlank("D")]];
  sheet.getRange(`E${totalRow}`).formulas = [[sumOrBlank("E")]];
  sheet.getRange(`F${totalRow}`).formulas = [[`=IF(COUNT(D${totalRow})=0,"",IF(D${totalRow}=0,"n.a.",IF(COUNT(E${totalRow})=0,"",E${totalRow}/D${totalRow})))`]];
  sheet.getRange(`G${totalRow}`).formulas = [[sumOrBlank("G")]];
  sheet.getRange(`H${totalRow}`).formulas = [[sumOrBlank("H")]];
  sheet.getRange(`I${totalRow}`).formulas = [[`=IF(COUNT(G${totalRow})=0,"",IF(G${totalRow}=0,"n.a.",IF(COUNT(H${totalRow})=0,"",H${totalRow}/G${totalRow})))`]];

  for (let w = 0; w < weeks.length; w += 1) {
    const base = 11 + w * 6;
    for (let offset = 0; offset < 6; offset += 1) {
      const col = columnName(base + offset);
      if (offset === 2 || offset === 5) {
        const planCol = columnName(base + offset - 2);
        const actualCol = columnName(base + offset - 1);
        sheet.getRange(`${col}${totalRow}`).formulas = [[`=IF(COUNT(${planCol}${totalRow})=0,"",IF(${planCol}${totalRow}=0,"n.a.",IF(COUNT(${actualCol}${totalRow})=0,"",${actualCol}${totalRow}/${planCol}${totalRow})))`]];
      } else {
        sheet.getRange(`${col}${totalRow}`).formulas = [[sumOrBlank(col)]];
      }
    }
  }
  const totalPublishPlanRefs = weeks.map((_, w) => `${columnName(11 + w * 6)}${totalRow}`);
  const totalPublishActualRefs = weeks.map((_, w) => [`${columnName(12 + w * 6)}${totalRow}`, `${columnName(12 + w * 6)}${totalRow}`, `${columnName(12 + w * 6)}${totalRow}`]);
  const totalAmountPlanRefs = weeks.map((_, w) => `${columnName(14 + w * 6)}${totalRow}`);
  const totalAmountActualRefs = weeks.map((_, w) => [`${columnName(15 + w * 6)}${totalRow}`, `${columnName(15 + w * 6)}${totalRow}`, `${columnName(15 + w * 6)}${totalRow}`]);
  const totalPublishStage = stageComponents(
    weeks,
    totalPublishPlanRefs,
    totalPublishActualRefs,
    controls,
  );
  const totalAmountStage = stageComponents(
    weeks,
    totalAmountPlanRefs,
    totalAmountActualRefs,
    controls,
  );
  const totalHelperCells = helperColsForRecord(taskRecords.length);
  sheet.getRange(`${totalHelperCells[0]}:${totalHelperCells[7]}`).formulas = [[
    `=${totalPublishStage.planCount}`,
    `=${totalPublishStage.planValue}`,
    `=${totalPublishStage.actualCount}`,
    `=${totalPublishStage.actualValue}`,
    `=${totalAmountStage.planCount}`,
    `=${totalAmountStage.planValue}`,
    `=${totalAmountStage.actualCount}`,
    `=${totalAmountStage.actualValue}`,
  ]];
  const totalPublishHelperStage = {
    stageRef: controls.stageRef,
    expectedCount: controls.weekRef,
    planCount: totalHelperCells[0],
    planValue: totalHelperCells[1],
    actualCount: totalHelperCells[2],
    actualValue: totalHelperCells[3],
  };
  const totalAmountHelperStage = {
    stageRef: controls.stageRef,
    expectedCount: controls.weekRef,
    planCount: totalHelperCells[4],
    planValue: totalHelperCells[5],
    actualCount: totalHelperCells[6],
    actualValue: totalHelperCells[7],
  };
  sheet.getRange(`J${totalRow}`).formulas = [[combinedStageStatusFormula(totalPublishHelperStage, totalAmountHelperStage)]];

  sheet.getRange(`A1:${lastCol}2`).format.fill = COLORS.navy;
  sheet.getRange(`A1:${lastCol}2`).format.font = { name: "Arial", size: 10, bold: true, color: COLORS.white };
  sheet.getRange(`A1:${lastCol}2`).format.horizontalAlignment = "center";
  sheet.getRange(`A1:${lastCol}${totalRow}`).format.borders = { preset: "all", style: "thin", color: COLORS.grid };
  sheet.getRange(`A${totalRow}:${lastCol}${totalRow}`).format.fill = COLORS.paleGray;
  sheet.getRange(`A${totalRow}:${lastCol}${totalRow}`).format.font = { name: "Arial", size: 10, bold: true, color: COLORS.text };
  sheet.getRange(`D${firstTaskRow}:E${totalRow}`).format.numberFormat = "#,##0";
  sheet.getRange(`F${firstTaskRow}:F${totalRow}`).format.numberFormat = "0.0%";
  sheet.getRange(`G${firstTaskRow}:H${totalRow}`).format.numberFormat = CURRENCY_FORMAT;
  for (const col of ["F", "I"]) {
    const rateRange = sheet.getRange(`${col}${firstTaskRow}:${col}${totalRow}`);
    rateRange.format.numberFormat = "0.0%";
    rateRange.conditionalFormats.add("dataBar", {
      color: COLORS.blue,
      thresholds: [{ type: "num", value: 0 }, { type: "num", value: 1 }],
      gradient: true,
    });
  }
  addStatusFormatting(sheet.getRange(`J${firstTaskRow}:J${totalRow}`));
  for (let w = 0; w < weeks.length; w += 1) {
    const base = 11 + w * 6;
    for (const offset of [0, 1]) {
      const col = columnName(base + offset);
      sheet.getRange(`${col}${firstTaskRow}:${col}${totalRow}`).format.numberFormat = "#,##0";
    }
    for (const offset of [3, 4]) {
      const col = columnName(base + offset);
      sheet.getRange(`${col}${firstTaskRow}:${col}${totalRow}`).format.numberFormat = CURRENCY_FORMAT;
    }
    for (const offset of [2, 5]) {
      const col = columnName(base + offset);
      const rateRange = sheet.getRange(`${col}${firstTaskRow}:${col}${totalRow}`);
      rateRange.format.numberFormat = "0.0%";
      rateRange.conditionalFormats.add("dataBar", {
        color: COLORS.blue,
        thresholds: [{ type: "num", value: 0 }, { type: "num", value: 1 }],
        gradient: true,
      });
    }
  }

  const stageCountCells = weeks.map((_, w) => [0, 1, 2].map((stageIndex) => `${columnName(w * 3 + stageIndex + 1)}${autoHelperRow}`));
  for (let w = 0; w < weeks.length; w += 1) {
    const detailBase = 9 + w * 7;
    for (let stageIndex = 0; stageIndex < 3; stageIndex += 1) {
      const detailCol = columnName(detailBase + [1, 3, 5][stageIndex]);
      const countTerms = taskRecords.map((record) => `COUNT(${indirectRef(record.sheetName, `${detailCol}3:${detailCol}6`)})`);
      sheet.getRange(stageCountCells[w][stageIndex]).formulas = [[`=SUM(${countTerms.join(",")})`]];
    }
  }

  let naturalWeekExpression = String(weeks.length);
  for (let w = weeks.length - 2; w >= 0; w -= 1) {
    naturalWeekExpression = `IF(TODAY()<=DATE(${year},${monthNumber},${weeks[w].endDay}),${w + 1},${naturalWeekExpression})`;
  }
  naturalWeekExpression = `IF(TODAY()<DATE(${year},${monthNumber},1),1,${naturalWeekExpression})`;
  const detectedWeekExpression = `MAX(1,${weeks.map((_, w) => `IF(AND((${naturalWeekExpression})>=${w + 1},SUM(${stageCountCells[w][0]}:${stageCountCells[w][2]})>0),${w + 1},0)`).join(",")})`;
  const stageForWeek = (w) => `IF(${stageCountCells[w][2]}>0,"整周",IF(${stageCountCells[w][1]}>0,"周尾",IF(${stageCountCells[w][0]}>0,"周中","待填写")))`;
  const detectedWeekCell = `${columnName(weeks.length * 3 + 1)}${autoHelperRow}`;
  let detectedStageExpression = stageForWeek(weeks.length - 1);
  for (let w = weeks.length - 2; w >= 0; w -= 1) {
    detectedStageExpression = `IF(${detectedWeekCell}=${w + 1},${stageForWeek(w)},${detectedStageExpression})`;
  }

  sheet.getRange(`A${controlRow}`).values = [[checkpoint.mode === "auto" ? "自动检查周" : "当前检查周"]];
  sheet.getRange(`C${controlRow}`).values = [[checkpoint.mode === "auto" ? "自动阶段" : "当前阶段"]];
  if (checkpoint.mode === "auto") {
    sheet.getRange(detectedWeekCell).formulas = [[`=${detectedWeekExpression}`]];
    sheet.getRange(`B${controlRow}`).formulas = [[`="W"&${detectedWeekCell}`]];
    sheet.getRange(`D${controlRow}`).formulas = [[`=${detectedStageExpression}`]];
  } else {
    sheet.getRange(`B${controlRow}`).values = [[checkpoint.week]];
    sheet.getRange(`D${controlRow}`).values = [[checkpoint.stage]];
    sheet.getRange(`B${controlRow}`).dataValidation = { rule: { type: "whole", operator: "between", formula1: 1, formula2: weeks.length } };
    sheet.getRange(`D${controlRow}`).dataValidation = { rule: { type: "list", values: ["周中", "周尾", "整周"] } };
  }
  sheet.getRange(`E${controlRow}`).values = [["当前应达比例"]];
  sheet.getRange(`F${controlRow}`).formulas = [[`=IF(D${controlRow}="周中",50%,IF(D${controlRow}="周尾",85%,IF(D${controlRow}="整周",100%,"")))`]];
  sheet.getRange(`A${controlRow}:F${controlRow}`).format.borders = { preset: "all", style: "thin", color: COLORS.grid };
  sheet.getRange(`A${controlRow}:F${controlRow}`).format.rowHeight = 28;
  for (const col of ["A", "C", "E"]) sheet.getRange(`${col}${controlRow}`).format.fill = COLORS.navy;
  sheet.getRange(`A${controlRow}:F${controlRow}`).format.horizontalAlignment = "center";
  for (const col of ["A", "C", "E"]) sheet.getRange(`${col}${controlRow}`).format.font = { name: "Arial", size: 10, bold: true, color: COLORS.white };
  for (const col of ["B", "D"]) {
    sheet.getRange(`${col}${controlRow}`).format.fill = checkpoint.mode === "auto" ? COLORS.paleBlue : COLORS.input;
    sheet.getRange(`${col}${controlRow}`).format.font = { name: "Arial", size: 10, bold: true, color: COLORS.text };
  }
  if (checkpoint.mode !== "auto") sheet.getRange(`B${controlRow}`).format.numberFormat = '"W"0';
  sheet.getRange(`F${controlRow}`).format.fill = COLORS.paleBlue;
  sheet.getRange(`F${controlRow}`).format.numberFormat = "0%";

  sheet.mergeCells(`A${noteRow}:J${noteRow}`);
  sheet.getRange(`A${noteRow}`).values = [["自动检查点：只在当天所属自然周及以前查找数据；同周按整周、周尾、周中的顺序识别最新已填写阶段，未来周误填不会提前推进。进度判断：已结束周按100%周计划计入，当前周按周中50%、周尾85%、整周100%计入；汇总取发布数与发布金额中的较弱项。"]];
  sheet.getRange(`A${noteRow}:J${noteRow}`).format.fill = COLORS.paleGray;
  sheet.getRange(`A${noteRow}:J${noteRow}`).format.font = { name: "Arial", size: 9, italic: true, color: "#595959" };

  const widths = { A: 12, B: 18, C: 15, D: 13, E: 13, F: 13, G: 15, H: 15, I: 13, J: 14 };
  for (const [col, width] of Object.entries(widths)) sheet.getRange(`${col}1:${col}${noteRow}`).format.columnWidth = width;
  for (let w = 0; w < weeks.length; w += 1) {
    const base = 11 + w * 6;
    const weekWidths = [12, 12, 11, 14, 14, 11];
    for (let offset = 0; offset < 6; offset += 1) {
      const col = columnName(base + offset);
      sheet.getRange(`${col}1:${col}${noteRow}`).format.columnWidth = weekWidths[offset];
    }
  }
  sheet.getRange(`A1:${lastCol}1`).format.rowHeight = 28;
  sheet.getRange(`A2:${lastCol}2`).format.rowHeight = 38;
  sheet.getRange(`A${firstTaskRow}:${lastCol}${totalRow}`).format.rowHeight = 28;
  sheet.getRange(`A${firstTaskRow}:${lastCol}${totalRow}`).format.verticalAlignment = "center";
  sheet.getRange(`D${firstTaskRow}:${lastCol}${totalRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`A${noteRow}:${lastCol}${noteRow}`).format.rowHeight = 42;
  sheet.getRange(`A${noteRow}:J${noteRow}`).format.wrapText = true;
  for (const row of helperRows) {
    sheet.getRange(`A${row}:${lastCol}${row}`).format.rowHeight = 2;
    sheet.getRange(`A${row}:${lastCol}${row}`).format.font = { name: "Arial", size: 1, color: COLORS.white };
    sheet.getRange(`A${row}:${lastCol}${row}`).format.fill = COLORS.white;
  }
  sheet.getRange(`A${autoHelperRow}:${lastCol}${autoHelperRow}`).format.rowHeight = 2;
  sheet.getRange(`A${autoHelperRow}:${lastCol}${autoHelperRow}`).format.font = { name: "Arial", size: 1, color: COLORS.white };
  sheet.getRange(`A${autoHelperRow}:${lastCol}${autoHelperRow}`).format.fill = COLORS.white;
  sheet.freezePanes.freezeRows(2);
  sheet.freezePanes.freezeColumns(10);
}

async function loadArtifactTool() {
  const resolver = createRequire(path.join(process.cwd(), "package.json"));
  let resolved;
  try {
    resolved = resolver.resolve("@oai/artifact-tool");
  } catch {
    throw new Error("Cannot resolve @oai/artifact-tool. Run from a temporary directory whose node_modules links to the workspace dependency node_modules directory.");
  }
  return import(pathToFileURL(resolved).href);
}

async function pathExists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const configPath = path.resolve(args.config);
  const config = JSON.parse(await fs.readFile(configPath, "utf8"));
  const month = String(config.month || "");
  const { year, monthNumber } = monthParts(month);
  const { weeks, daysInMonth } = buildWeeks(year, monthNumber);
  if (!Array.isArray(config.tasks) || config.tasks.length === 0) {
    throw new Error("tasks must contain at least one task.");
  }
  const tasks = config.tasks.map((task, i) => normalizeTask(task, i, weeks.length));
  const checkpoint = normalizeCheckpoint(config.checkpoint, weeks.length);
  const output = path.resolve(args.output || config.output || "");
  if (!path.isAbsolute(args.output || config.output || "")) throw new Error("output must be an absolute path.");
  if (path.extname(output).toLowerCase() !== ".xlsx") throw new Error("output must end with .xlsx.");
  if (!args.force && await pathExists(output)) {
    throw new Error(`Output already exists: ${output}. Use --force only when overwrite is authorized.`);
  }

  const { SpreadsheetFile, Workbook } = await loadArtifactTool();
  const workbook = Workbook.create();
  const summaryName = `${month}汇总`;
  const summarySheet = workbook.worksheets.add(summaryName);
  const totalRow = 3 + tasks.length;
  const controlRow = totalRow + 2;
  const autoHelperRow = controlRow + 3;
  const autoWeekCell = `${columnName(weeks.length * 3 + 1)}${autoHelperRow}`;
  const escapedSummaryName = summaryName.replace(/'/g, "''");
  const controls = {
    controlRow,
    weekRef: checkpoint.mode === "auto" ? `'${escapedSummaryName}'!$${autoWeekCell.replace(/\d+$/, "")}$${autoHelperRow}` : `'${escapedSummaryName}'!$B$${controlRow}`,
    stageRef: `'${escapedSummaryName}'!$D$${controlRow}`,
    ratioRef: `'${escapedSummaryName}'!$F$${controlRow}`,
  };
  const usedNames = new Set([summaryName]);
  const taskRecords = tasks.map((task) => {
    const sheetName = safeSheetName(task.name, usedNames);
    return { task, sheetName, sheet: workbook.worksheets.add(sheetName) };
  });

  for (const record of taskRecords) {
    buildTaskSheet(record.sheet, record.task, weeks, year, monthNumber, daysInMonth, controls);
  }
  buildSummarySheet(summarySheet, taskRecords, weeks, year, monthNumber, daysInMonth, checkpoint, controls);

  workbook.recalculate();
  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
    options: { useRegex: true, maxResults: 100 },
    summary: "generated dashboard formula error scan",
    maxChars: 12000,
  });
  const formulaErrors = errors.ndjson
    .trim()
    .split("\n")
    .filter(Boolean)
    .map((line) => JSON.parse(line))
    .filter((item) => item.kind === "match");
  if (formulaErrors.length > 0) {
    throw new Error(`Formula errors found before export:\n${formulaErrors.map((item) => JSON.stringify(item)).join("\n")}`);
  }

  await fs.mkdir(path.dirname(output), { recursive: true });
  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(output);
  await fs.rm(`${output}.inspect.ndjson`, { force: true });
  console.log(JSON.stringify({
    output,
    month,
    checkpoint: checkpoint.mode === "auto"
      ? { mode: "auto", note: "Current week and stage are calculated in the workbook after recalculation." }
      : checkpoint,
    week_count: weeks.length,
    sheets: [summaryName, ...taskRecords.map((record) => record.sheetName)],
    recalculation_required: true,
    recalculation_note: "Open and save in LibreOffice or Excel before delivery because summary formulas use INDIRECT for non-prefixed Chinese sheet names.",
  }, null, 2));
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
