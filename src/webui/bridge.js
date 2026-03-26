window.hostBridge = {
  _pendingHandlers: {},
  _requestCounter: 0,
  sendCommand: function (name, params, onResult) {
    var stamp = new Date().getTime();
    var requestId = 'r' + (++window.hostBridge._requestCounter);
    var key;
    var hash = 'cmd=' + encodeURIComponent(name) + '&rid=' + encodeURIComponent(requestId) + '&t=' + stamp;

    if (typeof params === 'function') {
      onResult = params;
      params = null;
    }

    if (typeof onResult === 'function') {
      window.hostBridge._pendingHandlers[requestId] = onResult;
    }

    if (params) {
      for (key in params) {
        if (params.hasOwnProperty(key)) {
          hash += '&' + encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
        }
      }
    }

    window.location.hash = hash;
    return requestId;
  },
  ping: function (onResult) {
    return window.hostBridge.sendCommand('ping', onResult);
  },
  pickRange: function (sheetName, onResult) {
    return window.hostBridge.sendCommand('pickRange', { sheet: sheetName || '' }, onResult);
  },
  getInitData: function (onResult) {
    return window.hostBridge.sendCommand('getInitData', onResult);
  },
  getSheetFields: function (sheetName, rangeText, onResult) {
    return window.hostBridge.sendCommand('getSheetFields', {
      sheet: sheetName || '',
      range: rangeText || ''
    }, onResult);
  },
  validateRange: function (sheetName, rangeText, onResult) {
    return window.hostBridge.sendCommand('validateRange', {
      sheet: sheetName || '',
      range: rangeText || ''
    }, onResult);
  },
  generateTemplate: function (payload, onResult) {
    return window.hostBridge.sendCommand('generateTemplate', payload, onResult);
  },
  writeData: function (payload, onResult) {
    return window.hostBridge.sendCommand('writeData', payload, onResult);
  },
  onProgress: function (payload) {
    var status = document.getElementById('status');
    var progressText = document.getElementById('progressText');
    var progressPercent = document.getElementById('progressPercent');
    var progressBar = document.getElementById('progressBar');
    var percent = 0;

    payload = payload || {};

    if (typeof payload.percent === 'number') {
      percent = payload.percent;
    } else if (!isNaN(Number(payload.percent))) {
      percent = Number(payload.percent);
    }

    if (percent < 0) percent = 0;
    if (percent > 100) percent = 100;

    if (status) {
      status.innerText = payload.message ? String(payload.message) : '';
    }
    if (progressText) {
      progressText.innerText = payload.message ? String(payload.message) : '执行中';
    }
    if (progressPercent) {
      progressPercent.innerText = percent + '%';
    }
    if (progressBar) {
      progressBar.style.width = percent + '%';
    }

    if (window.dispatchEvent && typeof window.CustomEvent === 'function') {
      window.dispatchEvent(new CustomEvent('host:progress', { detail: payload }));
    }
  },
  setOutput: function (text) {
  },
  setStatus: function (text) {
    var status = document.getElementById('status');
    if (status) {
      status.innerText = String(text || '');
    }
  },
  resolveCommand: function (requestId, resultText) {
    var handler = window.hostBridge._pendingHandlers[requestId];
    if (handler) {
      delete window.hostBridge._pendingHandlers[requestId];
      handler(resultText);
    }
  }
};

