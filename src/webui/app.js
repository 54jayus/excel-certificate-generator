(function () {
  var state = {
    init: null,
    fields: [],
    busy: false
  };

  var statusEl = document.getElementById('status');
  var progressBarEl = document.getElementById('progressBar');
  var progressTextEl = document.getElementById('progressText');
  var progressPercentEl = document.getElementById('progressPercent');

  var templateSheetEl = document.getElementById('templateSheet');
  var dataSheetEl = document.getElementById('dataSheet');
  var templateRangeEl = document.getElementById('templateRange');
  var pageFieldEl = document.getElementById('pageField');
  var perRowCountEl = document.getElementById('perRowCount');
  var generateCountEl = document.getElementById('generateCount');

  var pickRangeBtn = document.getElementById('pickRangeBtn');
  var generateBtn = document.getElementById('generateBtn');
  var writeBtn = document.getElementById('writeBtn');

  function tryParseJson(text) {
    try {
      return JSON.parse(text);
    } catch (error) {
      return null;
    }
  }

  function setStatus(text) {
    if (statusEl) {
      statusEl.innerText = text;
    }
  }

  function setProgress(text, percent) {
    var safePercent = Number(percent || 0);
    var trackWidth = 0;
    var pixelWidth = 0;

    if (isNaN(safePercent)) {
      safePercent = 0;
    }
    if (safePercent < 0) {
      safePercent = 0;
    }
    if (safePercent > 100) {
      safePercent = 100;
    }

    if (progressTextEl) {
      progressTextEl.innerText = text;
    }
    if (progressPercentEl) {
      progressPercentEl.innerText = safePercent + '%';
    }

    if (progressBarEl) {
      if (progressBarEl.parentNode) {
        trackWidth = progressBarEl.parentNode.clientWidth || progressBarEl.parentNode.offsetWidth || 0;
      }

      if (trackWidth > 0) {
        pixelWidth = Math.round(trackWidth * safePercent / 100);
        progressBarEl.style.width = pixelWidth + 'px';
      } else {
        progressBarEl.style.width = safePercent + '%';
      }

      progressBarEl.style.display = 'block';
      progressBarEl.style.zoom = '1';
    }
  }

  function setBusy(busy) {
    state.busy = busy;
    generateBtn.disabled = busy;
    writeBtn.disabled = busy;
    pickRangeBtn.disabled = busy;
  }

  function renderSelect(selectEl, items, placeholder) {
    var html = '';
    var i;
    if (placeholder) {
      html += '<option value="">' + placeholder + '</option>';
    }
    for (i = 0; i < items.length; i += 1) {
      html += '<option value="' + items[i] + '">' + items[i] + '</option>';
    }
    selectEl.innerHTML = html;
  }

  function collectForm() {
    return {
      templateSheet: templateSheetEl.value,
      dataSheet: dataSheetEl.value,
      templateRange: templateRangeEl.value.replace(/^\s+|\s+$/g, ''),
      pageField: pageFieldEl.value,
      perRowCount: Number(perRowCountEl.value || 0),
      generateCount: Number(generateCountEl.value || 0)
    };
  }

  function renderSummary() {
  }

  function setError(message) {
    setStatus('输入有误');
    setProgress(message ? message : '等待操作', 0);
    setBusy(false);
  }

  function validateGenerateForm(payload) {
    if (!payload.templateSheet) return '请选择模板 Sheet';
    if (!payload.templateRange) return '请选择模板区域';
    if (payload.perRowCount < 1) return '每行个数必须大于 0';
    if (payload.generateCount < 1) return '生成数量必须大于 0';
    return '';
  }

  function validateWriteForm(payload) {
    if (!payload.templateSheet) return '请选择模板 Sheet';
    if (!payload.dataSheet) return '请选择数据 Sheet';
    if (!payload.templateRange) return '请选择模板区域';
    if (!payload.pageField) return '请选择分页字段';
    if (payload.perRowCount < 1) return '每行个数必须大于 0';
    if (payload.generateCount < 1) return '生成数量必须大于 0';
    return '';
  }

  function loadTemplateFields() {
    var payload = collectForm();
    if (!payload.templateSheet || !payload.templateRange) {
      state.fields = [];
      renderSelect(pageFieldEl, [], '请先选择有效模板区域');
      renderSummary();
      return;
    }

    setStatus('加载分页字段中');
    window.hostBridge.getSheetFields(payload.templateSheet, payload.templateRange, function (resultText) {
      var data = tryParseJson(resultText);
      state.fields = data || [];
      renderSelect(pageFieldEl, ['不分页'].concat(state.fields), '请选择分页字段');
      pageFieldEl.value = '不分页';
      renderSummary();
      setStatus('分页字段已加载');
    });
  }

  function applyInitData(data) {
    var sheets = data.sheets || [];
    state.init = data;

    renderSelect(templateSheetEl, sheets, '请选择模板 Sheet');
    renderSelect(dataSheetEl, sheets, '请选择数据 Sheet');
    renderSelect(pageFieldEl, [], '请先选择有效模板区域');

    templateSheetEl.value = data.templateSheet || '';
    dataSheetEl.value = data.dataSheet || '';
    templateRangeEl.value = data.templateRange || '';
    perRowCountEl.value = data.perRowCount || 3;
    generateCountEl.value = data.generateCount || 1;

    renderSummary();
    setStatus('');
    setProgress('等待操作', 0);

    if (templateSheetEl.value && templateRangeEl.value) {
      validateCurrentRange(true);
    }
  }

  function loadInitData() {
    setStatus('加载工作簿配置中');
    window.hostBridge.getInitData(function (resultText) {
      var data = tryParseJson(resultText);
      if (data) {
        applyInitData(data);
      } else {
        setStatus('初始化失败');
      }
    });
  }

  function validateCurrentRange(shouldLoadFields) {
    var payload = collectForm();
    if (!payload.templateSheet || !payload.templateRange) {
      state.fields = [];
      renderSelect(pageFieldEl, [], '请先选择有效模板区域');
      renderSummary();
      return;
    }

    window.hostBridge.validateRange(payload.templateSheet, payload.templateRange, function (resultText) {
      var data = tryParseJson(resultText);
      if (!data) {
        return;
      }
      if (data.valid) {
        templateRangeEl.value = data.address;
        setStatus('模板区域有效');
        setProgress('区域校验通过', 16);
        renderSummary();
        if (shouldLoadFields !== false) {
          loadTemplateFields();
        }
      } else {
        state.fields = [];
        renderSelect(pageFieldEl, [], '请先选择有效模板区域');
        setError(data.message || '模板区域无效');
      }
    });
  }

  function handleActionResult(resultText, successStatus, successProgressText) {
    var data = tryParseJson(resultText);
    setBusy(false);
    if (data && data.success) {
      setStatus(successStatus);
      setProgress(successProgressText, 100);
    } else if (data) {
      setError(data.summaryMessage || '执行失败');
    } else {
      setStatus('返回结果不可解析');
    }
  }

  window.addEventListener('host:progress', function (event) {
    var payload = event.detail || {};
    if (payload.message) {
      setStatus(payload.message);
    }
    if (typeof payload.percent === 'number') {
      setProgress(payload.message || '执行中', payload.percent);
    }
  });

  templateSheetEl.onchange = function () {
    state.fields = [];
    renderSelect(pageFieldEl, [], '请先选择有效模板区域');
    renderSummary();
    validateCurrentRange();
  };

  dataSheetEl.onchange = function () {
    renderSummary();
  };

  templateRangeEl.onblur = function () {
    renderSummary();
    validateCurrentRange();
  };

  templateRangeEl.onkeyup = renderSummary;
  perRowCountEl.onchange = renderSummary;
  generateCountEl.onchange = renderSummary;
  pageFieldEl.onchange = renderSummary;

  pickRangeBtn.onclick = function () {
    setStatus('等待选择区域');
    window.hostBridge.pickRange(templateSheetEl.value, function (resultText) {
      templateRangeEl.value = resultText;
      renderSummary();
      validateCurrentRange();
    });
  };

  generateBtn.onclick = function () {
    var payload = collectForm();
    var error = validateGenerateForm(payload);
    if (error) {
      setError(error);
      return;
    }

    setBusy(true);
    setStatus('正在批量生成模板');
    setProgress('正在调用 VBA 生成模板', 35);

    window.hostBridge.generateTemplate({
      templateSheet: payload.templateSheet,
      templateRange: payload.templateRange,
      generateCount: String(payload.generateCount),
      perRowCount: String(payload.perRowCount)
    }, function (resultText) {
      handleActionResult(resultText, '批量生成模板已完成', '生成完成');
    });
  };

  writeBtn.onclick = function () {
    var payload = collectForm();
    var error = validateWriteForm(payload);
    if (error) {
      setError(error);
      return;
    }

    setBusy(true);
    setStatus('正在批量写入数据');
    setProgress('正在调用 VBA 写入数据', 35);

    window.hostBridge.writeData({
      dataSheet: payload.dataSheet,
      templateSheet: payload.templateSheet,
      templateRange: payload.templateRange,
      generateCount: String(payload.generateCount),
      perRowCount: String(payload.perRowCount),
      pageField: payload.pageField === '不分页' ? '' : payload.pageField
    }, function (resultText) {
      handleActionResult(resultText, '批量写入数据已完成', '写入完成');
    });
  };

  loadInitData();
})();
