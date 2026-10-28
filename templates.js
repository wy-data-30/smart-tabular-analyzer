// 场景模板：模板定义、字段映射表单、各模板的分析计算和结果渲染。

const TEMPLATE_DEFINITIONS = {
  generic: {
    name: "通用数据分析",
    fields: [
      { key: "metric", label: "分析指标字段", required: true, expected: "numeric" },
      { key: "group", label: "分组维度字段", required: true, expected: "category" },
      { key: "date", label: "时间字段（可选）", required: false, expected: "date" }
    ]
  },
  sales: {
    name: "销售数据分析",
    fields: [
      { key: "salesAmount", label: "销售额字段（可选）", required: false, expected: "numeric" },
      { key: "unitPrice", label: "单价字段（可选）", required: false, expected: "numeric" },
      { key: "quantity", label: "数量字段（可选）", required: false, expected: "numeric" },
      { key: "category", label: "品类字段", required: true, expected: "category" },
      { key: "region", label: "地区字段", required: true, expected: "category" },
      { key: "date", label: "日期字段（可选）", required: false, expected: "date" }
    ]
  },
  score: {
    name: "成绩数据分析",
    fields: [
      { key: "score", label: "成绩字段", required: true, expected: "numeric" },
      { key: "class", label: "班级字段（可选）", required: false, expected: "category" },
      { key: "subject", label: "科目字段（可选）", required: false, expected: "category" },
      { key: "examDate", label: "考试日期字段（可选）", required: false, expected: "date" }
    ]
  },
  usedGoods: {
    name: "二手商品价格分析",
    fields: [
      { key: "price", label: "价格字段", required: true, expected: "numeric" },
      { key: "item", label: "商品字段（可选）", required: false, expected: "category" },
      { key: "platform", label: "平台字段（可选）", required: false, expected: "category" },
      { key: "city", label: "城市字段（可选）", required: false, expected: "category" },
      { key: "condition", label: "成色字段（可选）", required: false, expected: "category" },
      { key: "date", label: "日期字段（可选）", required: false, expected: "date" }
    ]
  },
  survey: {
    name: "问卷调查分析",
    fields: [
      { key: "satisfaction", label: "满意度字段（可选）", required: false, expected: "numeric" },
      { key: "age", label: "年龄字段（可选）", required: false, expected: "numeric" },
      { key: "gender", label: "性别字段（可选）", required: false, expected: "category" },
      { key: "city", label: "城市字段（可选）", required: false, expected: "category" },
      { key: "choice", label: "选择题字段（可选）", required: false, expected: "category" }
    ]
  },
  behavior: {
    name: "用户行为分析",
    fields: [
      { key: "userId", label: "用户 ID 字段", required: true, expected: "any" },
      { key: "event", label: "事件字段", required: true, expected: "category" },
      { key: "channel", label: "渠道字段（可选）", required: false, expected: "category" },
      { key: "device", label: "设备字段（可选）", required: false, expected: "category" },
      { key: "date", label: "日期字段（可选）", required: false, expected: "date" }
    ]
  }
};

function getCurrentTemplate() {
  return TEMPLATE_DEFINITIONS[dom.scenarioTemplate.value] || TEMPLATE_DEFINITIONS.generic;
}

function renderTemplateMappingForm() {
  const template = getCurrentTemplate();
  dom.templateMappingForm.innerHTML = template.fields.map((field) => `
    <label>
      ${escapeHtml(field.label)}${field.required ? " *" : ""}
      <select id="templateField_${escapeHtml(field.key)}" data-template-field="${escapeHtml(field.key)}">
        ${buildTemplateFieldOptions(field)}
      </select>
    </label>
  `).join("");
}

function buildTemplateFieldOptions(fieldConfig) {
  const placeholder = fieldConfig.required ? "请选择字段" : "不选择";
  const options = getActiveProfiles()
    .filter((profile) => fieldConfig.expected === "any" || profile.typeKey === fieldConfig.expected)
    .filter((profile) => fieldConfig.expected !== "numeric" || fieldHasNumericData(profile.field))
    .filter((profile) => fieldConfig.expected !== "date" || fieldHasDateData(profile.field))
    .map((profile) => {
      return `<option value="${escapeHtml(profile.field)}">${escapeHtml(profile.field)}（${escapeHtml(profile.type)}）</option>`;
    }).join("");
  return `<option value="">${placeholder}</option>${options}`;
}

function resetTemplateResults(message) {
  clearTemplateCharts();
  dom.templateResults.classList.add("hidden");
  dom.templatePrompt.className = "analysis-prompt";
  dom.templatePrompt.textContent = message;
}

function clearTemplateCharts() {
  destroyChart("templateMainChart");
  destroyChart("templateSecondaryChart");
  destroyChart("templateTrendChart");
}

function getTemplateMappings(template) {
  return template.fields.reduce((mappings, field) => {
    const select = document.getElementById(`templateField_${field.key}`);
    mappings[field.key] = select ? select.value : "";
    return mappings;
  }, {});
}

function runTemplateAnalysis() {
  const templateKey = dom.scenarioTemplate.value;
  const template = getCurrentTemplate();
  const mappings = getTemplateMappings(template);
  const validationMessage = validateTemplateMappings(template, mappings);

  clearTemplateCharts();
  if (validationMessage) {
    dom.templateResults.classList.add("hidden");
    dom.templatePrompt.className = "analysis-prompt warning";
    dom.templatePrompt.textContent = validationMessage;
    return;
  }

  const result = {
    generic: analyzeGenericTemplate,
    sales: analyzeSalesTemplate,
    score: analyzeScoreTemplate,
    usedGoods: analyzeUsedGoodsTemplate,
    survey: analyzeSurveyTemplate,
    behavior: analyzeBehaviorTemplate
  }[templateKey](mappings);

  renderTemplateResult(result);
}

function validateTemplateMappings(template, mappings) {
  if (!getAnalysisRows().length) {
    return "当前筛选条件下没有可分析的记录，请调整或清除筛选条件。";
  }
  for (const field of template.fields) {
    const selected = mappings[field.key];
    if (field.required && !selected) {
      return `请为「${field.label.replace("（可选）", "")}」选择字段。`;
    }
    if (!selected) continue;
    const profile = getFieldProfile(selected);
    if (!profile || profile.typeKey === "ignore") {
      return `字段「${selected}」已被忽略或不再可用，请重新选择字段。`;
    }
    if (field.expected !== "any" && profile.typeKey !== field.expected) {
      return `字段「${selected}」当前配置为「${profile.type}」，与「${field.label}」要求的类型不匹配。`;
    }
    if (field.expected === "numeric" && !fieldHasNumericData(selected)) {
      return `字段「${selected}」无法解析出有效数值，请为「${field.label}」选择数值字段。`;
    }
    if (field.expected === "date" && !fieldHasDateData(selected)) {
      return `字段「${selected}」无法解析出有效日期，请为「${field.label}」选择日期字段或保持不选择。`;
    }
  }

  if (dom.scenarioTemplate.value === "survey") {
    const selectedCount = Object.values(mappings).filter(Boolean).length;
    if (selectedCount === 0) {
      return "问卷调查分析至少需要选择一个字段。";
    }
  }

  return "";
}

function renderTemplateResult(result) {
  dom.templatePrompt.className = "analysis-prompt";
  dom.templatePrompt.textContent = result.prompt;
  dom.templateResults.classList.remove("hidden");
  dom.templateSummaryCards.innerHTML = result.cards.map(([label, value]) => `
    <div class="analysis-summary-card">
      <div class="analysis-summary-label">${escapeHtml(label)}</div>
      <div class="analysis-summary-value" title="${escapeHtml(value)}">${escapeHtml(value)}</div>
    </div>
  `).join("");
  dom.templateInsights.innerHTML = result.insights.map((text) => `
    <div class="insight-item">${escapeHtml(text)}</div>
  `).join("");
  dom.templatePrimaryTable.innerHTML = buildSimpleTableHtml(result.table.headers, result.table.rows);

  renderTemplateChart("templateMainChart", dom.templateMainChartTitle, result.mainChart);
  renderTemplateChart("templateSecondaryChart", dom.templateSecondaryChartTitle, result.secondaryChart);
  if (result.trendChart) {
    dom.templateTrendCard.classList.remove("hidden");
    renderTemplateChart("templateTrendChart", dom.templateTrendChartTitle, result.trendChart);
  } else {
    dom.templateTrendCard.classList.add("hidden");
  }
}

function renderTemplateChart(chartId, titleElement, chartConfig) {
  const canvas = document.getElementById(chartId);
  titleElement.textContent = chartConfig.title;
  if (!window.Chart) {
    drawEmptyChart(canvas, "图表组件加载失败，请刷新页面重试");
    return;
  }
  if (!chartConfig.labels.length) {
    drawEmptyChart(canvas, "没有足够数据生成该图表");
    return;
  }

  state.charts[chartId] = new Chart(canvas, {
    type: chartConfig.type || "bar",
    data: {
      labels: chartConfig.labels,
      datasets: [{
        label: chartConfig.label,
        data: chartConfig.values,
        borderColor: chartConfig.color || CHART_THEME.primary,
        backgroundColor: chartConfig.background || chartConfig.color || CHART_THEME.primary,
        tension: 0.25,
        fill: chartConfig.type === "line"
      }]
    },
    options: chartOptions(
      numericAxisLabel(chartConfig.valueField, chartConfig.axisLabel || "值"),
      chartConfig.valueField
    )
  });
}

function analyzeGenericTemplate(mappings) {
  const grouped = groupNumericFieldByCategory(mappings.metric, mappings.group);
  const values = getNumericValues(mappings.metric);
  const stats = summarizeNumbers(values);
  const topBySum = grouped.find((item) => Number.isFinite(item.sum));
  const topByAvg = [...grouped]
    .sort((a, b) => compareNumbersDescending(a.avg, b.avg))[0];
  const trend = mappings.date ? buildTrendFromRows(mappings.date, (row) => toNumber(row[mappings.metric]), "sum") : null;
  const insights = [
    `本次通用分析使用「${mappings.metric}」作为数值指标、「${mappings.group}」作为分组维度，有效数值记录为 ${formatInteger(values.length)} 条。`,
    `「${mappings.metric}」的平均值为 ${formatFieldNumber(mappings.metric, stats.mean)}，中位数为 ${formatFieldNumber(mappings.metric, stats.median)}，最小值为 ${formatFieldNumber(mappings.metric, stats.min)}，最大值为 ${formatFieldNumber(mappings.metric, stats.max)}。`,
    topBySum ? `按「${mappings.group}」分组求和最高的是「${topBySum.name}」，求和值为 ${formatFieldNumber(mappings.metric, topBySum.sum)}。` : "所选字段没有形成可比较的分组结果。",
    topByAvg ? `按「${mappings.group}」分组平均值最高的是「${topByAvg.name}」，平均值为 ${formatFieldNumber(mappings.metric, topByAvg.avg)}。该结果只描述数据差异，不说明原因。` : "所选字段没有形成平均值比较结果。"
  ];
  if (trend) insights.push(buildTrendFactText(mappings.metric, mappings.date, trend, "求和值"));

  return {
    prompt: "已生成通用字段映射分析。",
    cards: [
      ["模板", "通用数据分析"],
      ["指标字段", mappings.metric],
      ["分组字段", mappings.group],
      ["有效记录", formatInteger(values.length)]
    ],
    insights,
    table: numericGroupTable(mappings.group, mappings.metric, grouped),
    mainChart: barChartConfig(`按 ${mappings.group} 的 ${mappings.metric} 求和 Top 10`, grouped.slice(0, 10).map((item) => item.name), grouped.slice(0, 10).map((item) => item.sum), `${mappings.metric} 求和`, "求和值", CHART_THEME.primary, mappings.metric),
    secondaryChart: histogramChartConfig(`${mappings.metric} 分布`, values, "记录数", mappings.metric),
    trendChart: trend ? lineChartConfig(`${mappings.metric} 时间趋势`, trend.labels, trend.values, `${mappings.metric} 求和`, "求和值", mappings.metric) : null
  };
}

function analyzeSalesTemplate(mappings) {
  const metricRows = buildSalesMetricRows(mappings);
  const categoryRows = metricRows
    ? groupMetricRowsByField(metricRows, mappings.category)
    : frequencyRows(mappings.category);
  const regionRows = metricRows
    ? groupMetricRowsByField(metricRows, mappings.region)
    : frequencyRows(mappings.region);
  const metricLabel = metricRows ? metricRows.label : "记录数";
  const metricFormatField = metricRows?.formatField || "";
  const categoryTop = metricRows
    ? categoryRows.find((item) => Number.isFinite(item.sum))
    : categoryRows[0];
  const regionTop = metricRows
    ? regionRows.find((item) => Number.isFinite(item.sum))
    : regionRows[0];
  const trend = mappings.date
    ? buildTrendFromRows(mappings.date, metricRows ? (row) => getSalesMetricValue(row, mappings) : null, metricRows ? "sum" : "count")
    : null;

  const insights = [
    `销售模板使用「${mappings.category}」和「${mappings.region}」进行分组，共分析 ${formatInteger(getAnalysisRows().length)} 行记录。`,
    metricRows
      ? `销售数值基于「${metricLabel}」计算，有效数值记录为 ${formatInteger(metricRows.length)} 条。`
      : "未选择可计算的销售数值字段，因此仅按记录数做分组统计。",
    categoryTop ? `按品类字段「${mappings.category}」比较，${metricRows ? "求和" : "记录数"}最高的分组是「${categoryTop.name}」，值为 ${metricRows ? formatFieldNumber(metricFormatField, categoryTop.sum) : formatNumber(categoryTop.count)}。` : "品类字段没有可统计的有效分组。",
    regionTop ? `按地区字段「${mappings.region}」比较，${metricRows ? "求和" : "记录数"}最高的分组是「${regionTop.name}」，值为 ${metricRows ? formatFieldNumber(metricFormatField, regionTop.sum) : formatNumber(regionTop.count)}。该结果只反映 CSV 中的数据事实。` : "地区字段没有可统计的有效分组。"
  ];
  if (trend) insights.push(buildTrendFactText(metricLabel, mappings.date, trend, metricRows ? "求和值" : "记录数"));

  return {
    prompt: "已生成销售数据模板分析。",
    cards: [
      ["模板", "销售数据分析"],
      ["销售指标", metricLabel],
      ["品类字段", mappings.category],
      ["地区字段", mappings.region]
    ],
    insights,
    table: metricRows
      ? numericGroupTable(mappings.category, metricLabel, categoryRows, metricFormatField)
      : frequencyTable(mappings.category, categoryRows),
    mainChart: barChartConfig(`按 ${mappings.category} 的 ${metricLabel} Top 10`, topLabels(categoryRows), topValues(categoryRows, metricRows ? "sum" : "count"), metricRows ? `${metricLabel} 求和` : "记录数", metricRows ? "求和值" : "记录数", CHART_THEME.primary, metricFormatField),
    secondaryChart: barChartConfig(`按 ${mappings.region} 的 ${metricLabel} Top 10`, topLabels(regionRows), topValues(regionRows, metricRows ? "sum" : "count"), metricRows ? `${metricLabel} 求和` : "记录数", metricRows ? "求和值" : "记录数", CHART_THEME.secondary, metricFormatField),
    trendChart: trend ? lineChartConfig(`${metricLabel} 时间趋势`, trend.labels, trend.values, metricRows ? `${metricLabel} 求和` : "记录数", metricRows ? "求和值" : "记录数", metricFormatField) : null
  };
}

function analyzeScoreTemplate(mappings) {
  const values = getNumericValues(mappings.score);
  const stats = summarizeNumbers(values);
  const dimensionResults = [...new Set([mappings.class, mappings.subject].filter(Boolean))]
    .map((field) => ({ field, rows: groupNumericFieldByCategory(mappings.score, field) }));
  const primaryDimension = dimensionResults[0];
  const secondaryDimension = dimensionResults[1];
  const trend = mappings.examDate ? buildTrendFromRows(mappings.examDate, (row) => toNumber(row[mappings.score]), "avg") : null;
  const insights = [
    `成绩模板使用「${mappings.score}」作为成绩字段，有效成绩记录为 ${formatInteger(values.length)} 条。`,
    `成绩平均值为 ${formatFieldNumber(mappings.score, stats.mean)}，中位数为 ${formatFieldNumber(mappings.score, stats.median)}，最低值为 ${formatFieldNumber(mappings.score, stats.min)}，最高值为 ${formatFieldNumber(mappings.score, stats.max)}。`
  ];
  if (!dimensionResults.length) {
    insights.push("未选择班级或科目字段，因此未生成分组平均值比较。");
  } else {
    dimensionResults.forEach(({ field, rows }) => {
      const sorted = [...rows].sort((a, b) => b.avg - a.avg);
      const topAvg = sorted[0];
      const lowAvg = sorted[sorted.length - 1];
      insights.push(topAvg && lowAvg
        ? `按「${field}」分组时，平均值最高的是「${topAvg.name}」(${formatFieldNumber(mappings.score, topAvg.avg)})，平均值最低的是「${lowAvg.name}」(${formatFieldNumber(mappings.score, lowAvg.avg)})。这里只描述字段结果，不评价教学或个体原因。`
        : `字段「${field}」没有形成可比较的成绩分组。`);
    });
  }
  if (trend) insights.push(buildTrendFactText(mappings.score, mappings.examDate, trend, "平均值"));

  return {
    prompt: "已生成成绩数据模板分析。",
    cards: [
      ["模板", "成绩数据分析"],
      ["成绩字段", mappings.score],
      ["有效成绩", formatInteger(values.length)],
      ["平均值", formatFieldNumber(mappings.score, stats.mean)]
    ],
    insights,
    table: dimensionResults.length ? multiDimensionNumericTable(mappings.score, dimensionResults) : simpleStatsTable(mappings.score, stats, values.length),
    mainChart: primaryDimension
      ? barChartConfig(`按 ${primaryDimension.field} 的成绩平均值 Top 10`, topLabels(primaryDimension.rows, "avg"), topValues(primaryDimension.rows, "avg"), "成绩平均值", "平均值", CHART_THEME.primary, mappings.score)
      : histogramChartConfig(`${mappings.score} 分布`, values, "记录数", mappings.score),
    secondaryChart: secondaryDimension
      ? barChartConfig(`按 ${secondaryDimension.field} 的成绩平均值 Top 10`, topLabels(secondaryDimension.rows, "avg"), topValues(secondaryDimension.rows, "avg"), "成绩平均值", "平均值", CHART_THEME.secondary, mappings.score)
      : histogramChartConfig(`${mappings.score} 分布`, values, "记录数", mappings.score),
    trendChart: trend ? lineChartConfig(`${mappings.score} 时间趋势`, trend.labels, trend.values, "成绩平均值", "平均值", mappings.score) : null
  };
}

function analyzeUsedGoodsTemplate(mappings) {
  const values = getNumericValues(mappings.price);
  const stats = summarizeNumbers(values);
  const dimensionResults = [...new Set(["item", "platform", "city", "condition"].map((key) => mappings[key]).filter(Boolean))]
    .map((field) => ({ field, rows: groupNumericFieldByCategory(mappings.price, field) }));
  const primaryDimension = dimensionResults[0];
  const secondaryDimension = dimensionResults[1];
  const trend = mappings.date ? buildTrendFromRows(mappings.date, (row) => toNumber(row[mappings.price]), "avg") : null;
  const insights = [
    `二手商品价格模板使用「${mappings.price}」作为价格字段，有效价格记录为 ${formatInteger(values.length)} 条。`,
    `价格平均值为 ${formatFieldNumber(mappings.price, stats.mean)}，中位数为 ${formatFieldNumber(mappings.price, stats.median)}，最低值为 ${formatFieldNumber(mappings.price, stats.min)}，最高值为 ${formatFieldNumber(mappings.price, stats.max)}。`
  ];
  if (!dimensionResults.length) {
    insights.push("未选择商品、平台、城市或成色字段，因此仅生成价格整体统计。");
  } else {
    dimensionResults.forEach(({ field, rows }) => {
      const topAvg = [...rows].sort((a, b) => b.avg - a.avg)[0];
      insights.push(topAvg
        ? `按「${field}」比较，平均价格最高的分组是「${topAvg.name}」，平均值为 ${formatFieldNumber(mappings.price, topAvg.avg)}。该结果只说明样本中的价格差异，不代表市场整体水平。`
        : `字段「${field}」没有形成可比较的价格分组。`);
    });
  }
  if (trend) insights.push(buildTrendFactText(mappings.price, mappings.date, trend, "平均值"));

  return {
    prompt: "已生成二手商品价格模板分析。",
    cards: [
      ["模板", "二手商品价格分析"],
      ["价格字段", mappings.price],
      ["比较维度", dimensionResults.length ? formatInteger(dimensionResults.length) : "未选择"],
      ["价格中位数", formatFieldNumber(mappings.price, stats.median)]
    ],
    insights,
    table: dimensionResults.length ? multiDimensionNumericTable(mappings.price, dimensionResults) : simpleStatsTable(mappings.price, stats, values.length),
    mainChart: primaryDimension
      ? barChartConfig(`按 ${primaryDimension.field} 的平均价格 Top 10`, topLabels(primaryDimension.rows, "avg"), topValues(primaryDimension.rows, "avg"), "平均价格", "平均值", CHART_THEME.primary, mappings.price)
      : histogramChartConfig(`${mappings.price} 分布`, values, "记录数", mappings.price),
    secondaryChart: secondaryDimension
      ? barChartConfig(`按 ${secondaryDimension.field} 的平均价格 Top 10`, topLabels(secondaryDimension.rows, "avg"), topValues(secondaryDimension.rows, "avg"), "平均价格", "平均值", CHART_THEME.secondary, mappings.price)
      : histogramChartConfig(`${mappings.price} 分布`, values, "记录数", mappings.price),
    trendChart: trend ? lineChartConfig(`${mappings.price} 时间趋势`, trend.labels, trend.values, "平均价格", "平均值", mappings.price) : null
  };
}

function analyzeSurveyTemplate(mappings) {
  const selectedFields = Object.entries(mappings).filter(([, value]) => value);
  const numericSummaries = [];
  ["satisfaction", "age"].forEach((key) => {
    if (!mappings[key]) return;
    const values = getNumericValues(mappings[key]);
    if (values.length) numericSummaries.push({ key, field: mappings[key], stats: summarizeNumbers(values), count: values.length });
  });
  const categoricalFields = ["gender", "city", "choice"].map((key) => mappings[key]).filter(Boolean);
  const primaryCategory = categoricalFields[0] || "";
  const primaryRows = primaryCategory ? frequencyRows(primaryCategory) : [];
  const chartValues = numericSummaries[0] ? getNumericValues(numericSummaries[0].field) : [];
  const insights = [
    `问卷模板共选择 ${formatInteger(selectedFields.length)} 个字段，当前分析记录数为 ${formatInteger(getAnalysisRows().length)}。`,
    numericSummaries[0]
      ? `数值字段「${numericSummaries[0].field}」的平均值为 ${formatFieldNumber(numericSummaries[0].field, numericSummaries[0].stats.mean)}，中位数为 ${formatFieldNumber(numericSummaries[0].field, numericSummaries[0].stats.median)}。`
      : "未选择可解析的数值字段，因此未生成满意度或年龄的数值统计。",
    primaryRows[0]
      ? `分类字段「${primaryCategory}」中出现最多的类别是「${primaryRows[0].name}」，占比 ${formatPercent(primaryRows[0].ratio)}。该结果只描述样本分布。`
      : "未选择分类字段，因此未生成选择题或人口属性频数统计。"
  ];

  return {
    prompt: "已生成问卷调查模板分析。",
    cards: [
      ["模板", "问卷调查分析"],
      ["已选字段", formatInteger(selectedFields.length)],
      ["记录数", formatInteger(getAnalysisRows().length)],
      ["主分类字段", primaryCategory || "未选择"]
    ],
    insights,
    table: surveyTable(numericSummaries, categoricalFields),
    mainChart: primaryCategory
      ? barChartConfig(`${primaryCategory} Top 10`, topLabels(primaryRows), topValues(primaryRows, "count"), "频数", "频数", CHART_THEME.primary)
      : histogramChartConfig(numericSummaries[0]?.field || "数值字段", chartValues, "记录数", numericSummaries[0]?.field || ""),
    secondaryChart: numericSummaries[0]
      ? histogramChartConfig(`${numericSummaries[0].field} 分布`, chartValues, "记录数", numericSummaries[0].field)
      : barChartConfig("分类字段 Top 10", topLabels(primaryRows), topValues(primaryRows, "count"), "频数", "频数", CHART_THEME.secondary),
    trendChart: null
  };
}

function analyzeBehaviorTemplate(mappings) {
  const eventRows = frequencyRows(mappings.event);
  const userCounts = countNonMissingValues(mappings.userId);
  const userIdKeys = getConfirmedPlaceholderKeys(mappings.userId);
  const uniqueUsers = new Set(getAnalysisRows().map((row) => analysisValue(row[mappings.userId], userIdKeys)).filter(Boolean)).size;
  const avgEventsPerUser = uniqueUsers ? userCounts / uniqueUsers : 0;
  const secondaryResults = [...new Set([mappings.channel, mappings.device].filter(Boolean))]
    .map((field) => ({ field, rows: frequencyRows(field) }));
  const primarySecondary = secondaryResults[0];
  const secondaryRows = primarySecondary ? primarySecondary.rows : topUserRows(mappings.userId);
  const trend = mappings.date ? buildTrendFromRows(mappings.date, null, "count") : null;
  const insights = [
    `用户行为模板使用「${mappings.userId}」作为用户 ID、「${mappings.event}」作为事件字段，有效事件记录为 ${formatInteger(userCounts)} 条。`,
    `唯一用户数为 ${formatInteger(uniqueUsers)}，平均每个用户对应 ${formatNumber(avgEventsPerUser)} 条事件记录。`,
    eventRows[0] ? `事件字段中出现最多的是「${eventRows[0].name}」，出现 ${formatInteger(eventRows[0].count)} 次，占比 ${formatPercent(eventRows[0].ratio)}。该结果不推断用户动机。` : "事件字段没有可统计的有效类别。"
  ];
  secondaryResults.forEach(({ field, rows }) => {
    if (rows[0]) insights.push(`辅助维度「${field}」中记录数最高的是「${rows[0].name}」，出现 ${formatInteger(rows[0].count)} 次。`);
  });
  if (trend) insights.push(buildTrendFactText("事件记录数", mappings.date, trend, "记录数"));

  return {
    prompt: "已生成用户行为模板分析。",
    cards: [
      ["模板", "用户行为分析"],
      ["唯一用户", formatInteger(uniqueUsers)],
      ["事件记录", formatInteger(userCounts)],
      ["人均事件", formatNumber(avgEventsPerUser)]
    ],
    insights,
    table: multiFieldFrequencyTable([{ field: mappings.event, rows: eventRows }, ...secondaryResults]),
    mainChart: barChartConfig(`${mappings.event} Top 10`, topLabels(eventRows), topValues(eventRows, "count"), "事件数", "记录数", CHART_THEME.primary),
    secondaryChart: barChartConfig(`${primarySecondary?.field || mappings.userId} Top 10`, topLabels(secondaryRows), topValues(secondaryRows, "count"), "记录数", "记录数", CHART_THEME.secondary),
    trendChart: trend ? lineChartConfig("事件记录时间趋势", trend.labels, trend.values, "记录数", "记录数") : null
  };
}

function buildSalesMetricRows(mappings) {
  let label = "";
  let formatField = "";
  if (mappings.salesAmount) {
    label = mappings.salesAmount;
    formatField = mappings.salesAmount;
  }
  else if (mappings.unitPrice && mappings.quantity) label = `${mappings.unitPrice} × ${mappings.quantity}`;
  else if (mappings.unitPrice) {
    label = mappings.unitPrice;
    formatField = mappings.unitPrice;
  }
  else if (mappings.quantity) {
    label = mappings.quantity;
    formatField = mappings.quantity;
  }
  else return null;

  const rows = getAnalysisRows()
    .map((row) => ({ row, value: getSalesMetricValue(row, mappings) }))
    .filter((item) => item.value !== null);
  rows.label = label;
  rows.formatField = formatField;
  return rows;
}

function getSalesMetricValue(row, mappings) {
  if (mappings.salesAmount) return toNumber(row[mappings.salesAmount]);
  if (mappings.unitPrice && mappings.quantity) {
    const price = toNumber(row[mappings.unitPrice]);
    const quantity = toNumber(row[mappings.quantity]);
    if (price === null || quantity === null) return null;
    const product = price * quantity;
    return Number.isFinite(product) ? product : null;
  }
  if (mappings.unitPrice) return toNumber(row[mappings.unitPrice]);
  if (mappings.quantity) return toNumber(row[mappings.quantity]);
  return null;
}

function buildTrendFactText(metricLabel, dateField, trend, valueLabel) {
  if (!trend.labels.length) return `已选择日期字段「${dateField}」，但没有足够的有效日期记录生成趋势。`;
  if (trend.labels.length === 1) return `日期字段「${dateField}」只有 1 个有效时间点，无法判断变化方向。`;
  const first = trend.values[0];
  const last = trend.values[trend.values.length - 1];
  if (!Number.isFinite(first) || !Number.isFinite(last)) {
    return `按日期字段「${dateField}」观察时，聚合结果超出浏览器可表示的有限数值范围，无法可靠判断变化方向。`;
  }
  const direction = last > first ? "上升" : last < first ? "下降" : "基本持平";
  return `按日期字段「${dateField}」观察，「${metricLabel}」${valueLabel}从 ${trend.labels[0]} 到 ${trend.labels[trend.labels.length - 1]} 呈${direction}变化；这只是时间序列数值变化，不说明原因。`;
}

function numericGroupTable(groupField, metricField, rows, formatField = metricField) {
  return {
    headers: [groupField, "记录数", `${metricField} 求和`, `${metricField} 平均值`, "最小值", "最大值"],
    rows: rows.map((row) => [
      row.name,
      formatInteger(row.count),
      formatFieldNumber(formatField, row.sum),
      formatFieldNumber(formatField, row.avg),
      formatFieldNumber(formatField, row.min),
      formatFieldNumber(formatField, row.max)
    ])
  };
}

function multiDimensionNumericTable(metricField, dimensionResults) {
  return {
    headers: ["维度字段", "分组值", "记录数", `${metricField} 求和`, `${metricField} 平均值`, "最小值", "最大值"],
    rows: dimensionResults.flatMap(({ field, rows }) => rows.map((row) => [
      field,
      row.name,
      formatInteger(row.count),
      formatFieldNumber(metricField, row.sum),
      formatFieldNumber(metricField, row.avg),
      formatFieldNumber(metricField, row.min),
      formatFieldNumber(metricField, row.max)
    ]))
  };
}

function frequencyTable(field, rows) {
  return {
    headers: [field, "记录数", "占比"],
    rows: rows.map((row) => [row.name, formatInteger(row.count), formatPercent(row.ratio)])
  };
}

function multiFieldFrequencyTable(fieldResults) {
  return {
    headers: ["字段", "类别", "记录数", "占比"],
    rows: fieldResults.flatMap(({ field, rows }) => rows.map((row) => [
      field,
      row.name,
      formatInteger(row.count),
      formatPercent(row.ratio)
    ]))
  };
}

function simpleStatsTable(field, stats, count) {
  return {
    headers: ["字段", "有效数量", "平均值", "中位数", "最小值", "最大值", "标准差"],
    rows: [[field, formatInteger(count), formatFieldNumber(field, stats.mean), formatFieldNumber(field, stats.median), formatFieldNumber(field, stats.min), formatFieldNumber(field, stats.max), formatFieldNumber(field, stats.std)]]
  };
}

function surveyTable(numericSummaries, categoricalFields) {
  const rows = [];
  numericSummaries.forEach((item) => {
    rows.push([item.field, "数值字段", `有效 ${formatInteger(item.count)} 条，平均值 ${formatFieldNumber(item.field, item.stats.mean)}，中位数 ${formatFieldNumber(item.field, item.stats.median)}`]);
  });
  categoricalFields.forEach((field) => {
    const top = frequencyRows(field)[0];
    rows.push([field, "分类字段", top ? `Top 类别「${top.name}」，占比 ${formatPercent(top.ratio)}` : "没有有效类别"]);
  });
  return {
    headers: ["字段", "类型", "结果摘要"],
    rows
  };
}

function buildSimpleTableHtml(headers, rows) {
  return `
    <thead>
      <tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr>
    </thead>
    <tbody>
      ${rows.map((row) => `
        <tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>
      `).join("")}
    </tbody>
  `;
}

function topUserRows(field) {
  return frequencyRows(field, 10);
}

function topLabels(rows, valueKey = "sum") {
  return [...rows]
    .sort((a, b) => compareNumbersDescending(a[valueKey], b[valueKey]))
    .slice(0, 10)
    .map((item) => item.name);
}

function topValues(rows, valueKey) {
  return [...rows]
    .sort((a, b) => compareNumbersDescending(a[valueKey], b[valueKey]))
    .slice(0, 10)
    .map((item) => Number.isFinite(item[valueKey]) ? item[valueKey] : null);
}

function barChartConfig(title, labels, values, label, axisLabel, color, valueField = "") {
  return { type: "bar", title, labels, values, label, axisLabel, color, valueField };
}

function lineChartConfig(title, labels, values, label, axisLabel, valueField = "") {
  return {
    type: "line",
    title,
    labels,
    values,
    label,
    axisLabel,
    valueField,
    color: CHART_THEME.secondary,
    background: CHART_THEME.secondaryFill
  };
}

function histogramChartConfig(title, values, axisLabel, sourceField = "") {
  const bins = buildHistogram(values, 10, sourceField);
  return {
    type: "bar",
    title,
    labels: bins.map((bin) => bin.label),
    values: bins.map((bin) => bin.count),
    label: title,
    axisLabel,
    color: CHART_THEME.warning
  };
}
