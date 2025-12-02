(() => {
  const ready = (fn) => {
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', fn);
    } else {
      fn();
    }
  };

  ready(() => {
    const refs = {
      fileInput: document.getElementById('files'),
      fileList: document.getElementById('file-list'),
      fileSelectAllBtn: document.getElementById('files-select-all'),
      fileDeleteBtn: document.getElementById('files-delete-selected'),
      fileClearBtn: document.getElementById('files-clear'),
      statusBox: document.getElementById('status'),
      downloadBox: document.getElementById('download'),
      mergeBtn: document.getElementById('merge-btn'),
      resetBtn: document.getElementById('reset-btn'),
      inspectBtn: document.getElementById('inspect-btn'),
      columnsBox: document.getElementById('columns-box'),
      previewArea: document.getElementById('preview-area'),
      colSelectAll: document.getElementById('col-select-all'),
      colUnselectAll: document.getElementById('col-unselect-all'),
      colInvert: document.getElementById('col-invert'),
      mappingPanel: document.getElementById('mapping-panel'),
      mappingToggle: document.getElementById('mapping-toggle'),
      mappingInput: document.getElementById('mapping-input'),
      mappingSave: document.getElementById('mapping-save'),
      mappingStatus: document.getElementById('mapping-status'),
      cleanupLogs: document.getElementById('cleanup-logs'),
      cleanupTemp: document.getElementById('cleanup-temp'),
      logBox: document.getElementById('log-box'),
    };

    if (!refs.fileInput || !refs.fileList || !refs.inspectBtn) {
      console.error('初始化失败，DOM 元素缺失');
      return;
    }

    let filesState = [];
    let fileId = 1;
    let lastPreviews = [];
    let mappingLoaded = false;

    const log = (msg) => {
      console.log(msg);
      if (!refs.logBox) return;
      const line = document.createElement('div');
      line.textContent = `[${new Date().toLocaleTimeString()}] ${msg}`;
      refs.logBox.appendChild(line);
      refs.logBox.scrollTop = refs.logBox.scrollHeight;
    };

    const setStatus = (text, isError = false) => {
      if (!refs.statusBox) return;
      refs.statusBox.textContent = text;
      refs.statusBox.style.display = 'block';
      refs.statusBox.classList.toggle('error', !!isError);
    };

    const renderFiles = () => {
      if (!filesState.length) {
        refs.fileList.textContent = '尚未选择文件。';
        return;
      }
      refs.fileList.innerHTML = filesState.map(f => {
        return `<label style="display:flex; gap:8px; align-items:center; margin-bottom:4px;">
          <input type="checkbox" class="file-item" data-id="${f.id}">
          <span>${f.name} — ${(f.size/1024).toFixed(1)} KB</span>
        </label>`;
      }).join('');
      refs.fileList.querySelectorAll('.file-item').forEach(cb => {
        cb.addEventListener('change', () => showPreviewForSelection());
      });
    };

    const renderColumns = (columns) => {
      if (!columns || !columns.length) {
        refs.columnsBox.textContent = '未获取到列信息。';
        return;
      }
      refs.columnsBox.innerHTML = columns.map(col => {
        const disabled = col.is_meta ? 'disabled' : '';
        const title = (col.sources || []).join(', ');
        return `<label style="display:flex; gap:8px; align-items:center; margin-bottom:6px;">
          <input type="checkbox" class="col-item" value="${col.name}" ${disabled}>
          <span>${col.name}</span>
          <small style="color:var(--muted);">(${title})${col.is_meta ? ' - 保留' : ''}</small>
        </label>`;
      }).join('');
    };

    const renderPreview = (previews) => {
      lastPreviews = previews || [];
      if (!lastPreviews.length) {
        refs.previewArea.textContent = '未获取到预览数据。';
        return;
      }
      refs.previewArea.innerHTML = lastPreviews.map(p => {
        const rows = p.rows || [];
        const header = (p.columns || []).join(' | ');
        const body = rows.map(r => Object.values(r).join(' | ')).join(' | ');
        return `<div style="margin-bottom:12px;"><strong>${p.file} / ${p.sheet}</strong><pre style="margin:6px 0; white-space:pre-wrap;">${header}\n${body}</pre></div>`;
      }).join('');
    };

    const collectExcludedColumns = () => {
      return Array.from(document.querySelectorAll('.col-item:checked')).map(cb => cb.value);
    };

    const showPreviewForSelection = () => {
      const selectedIds = new Set(Array.from(document.querySelectorAll('.file-item:checked')).map(cb => Number(cb.dataset.id)));
      if (!lastPreviews.length) return;
      if (!selectedIds.size) {
        renderPreview(lastPreviews);
        return;
      }
      const selectedNames = new Set(filesState.filter(f => selectedIds.has(f.id)).map(f => f.name));
      const filtered = lastPreviews.filter(p => selectedNames.has(p.file));
      renderPreview(filtered);
    };

    refs.fileInput.addEventListener('change', (e) => {
      const files = Array.from(e.target.files || []);
      files.forEach(f => {
        filesState.push({ id: fileId++, name: f.name, size: f.size, file: f });
        log(`已添加文件: ${f.name}`);
      });
      refs.fileInput.value = '';
      renderFiles();
    });

    refs.fileSelectAllBtn?.addEventListener('click', () => {
      document.querySelectorAll('.file-item').forEach(cb => cb.checked = true);
      showPreviewForSelection();
    });

    refs.fileDeleteBtn?.addEventListener('click', () => {
      const selected = new Set(Array.from(document.querySelectorAll('.file-item:checked')).map(cb => Number(cb.dataset.id)));
      filesState = filesState.filter(f => !selected.has(f.id));
      renderFiles();
      showPreviewForSelection();
      log('已删除选中文件');
    });

    refs.fileClearBtn?.addEventListener('click', () => {
      filesState = [];
      renderFiles();
      refs.previewArea.textContent = '预览区：点击“分析文件”后展示前 5 行。';
      log('已清空文件列表');
    });

    refs.resetBtn?.addEventListener('click', () => {
      refs.fileInput.value = '';
      filesState = [];
      renderFiles();
      refs.statusBox.style.display = 'none';
      refs.downloadBox.style.display = 'none';
      document.getElementById('dedup_keys').value = '';
      document.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = cb.id === 'normalize');
      refs.columnsBox.innerHTML = '请先点击“分析文件”加载列信息。';
      refs.previewArea.innerHTML = '预览区：点击“分析文件”后展示前 5 行。';
      lastPreviews = [];
      log('已重置表单');
    });

    refs.inspectBtn.addEventListener('click', async () => {
      const files = [...filesState];
      if (!files.length) {
        setStatus('请先选择文件。', true);
        return;
      }
      refs.columnsBox.textContent = '列信息加载中...';
      refs.previewArea.textContent = '预览加载中...';
      log(`开始分析，文件数：${files.length}`);

      const formData = new FormData();
      files.forEach(f => formData.append('files', f.file, f.name));
      if (document.getElementById('normalize').checked) formData.append('normalize_columns', 'on');
      if (document.getElementById('fuzzy').checked) formData.append('enable_fuzzy', 'on');

      setStatus('正在分析文件...');

      try {
        const res = await fetch('/inspect', { method: 'POST', body: formData, credentials: 'same-origin' });
        if (!res.ok) {
          setStatus(`分析失败（${res.status}）`, true);
          refs.previewArea.textContent = '预览失败';
          log(`分析失败，HTTP ${res.status}`);
          return;
        }
        const data = await res.json();
        if (!data.ok) {
          setStatus(data.error || '分析失败', true);
          refs.previewArea.textContent = '预览失败';
          log(`分析失败：${data.error || '未知错误'}`);
          return;
        }
        renderColumns(data.columns || []);
        renderPreview(data.previews || []);
        showPreviewForSelection();
        setStatus('分析完成，可选择要删除的列后开始合并。');
        refs.downloadBox.style.display = 'none';
        log('分析完成');
      } catch (err) {
        setStatus('分析请求失败，请稍后重试。', true);
        refs.previewArea.textContent = '预览失败';
        console.error(err);
        log('分析请求失败');
      }
    });

    refs.colSelectAll?.addEventListener('click', () => {
      document.querySelectorAll('.col-item:not(:disabled)').forEach(cb => cb.checked = true);
    });
    refs.colUnselectAll?.addEventListener('click', () => {
      document.querySelectorAll('.col-item').forEach(cb => cb.checked = false);
    });
    refs.colInvert?.addEventListener('click', () => {
      document.querySelectorAll('.col-item:not(:disabled)').forEach(cb => cb.checked = !cb.checked);
    });

    const loadMapping = async () => {
      try {
        const res = await fetch('/mapping');
        const data = await res.json();
        if (data.ok) {
          refs.mappingInput.value = JSON.stringify(data.mappings, null, 2);
          refs.mappingStatus.textContent = '映射已加载。';
        } else {
          refs.mappingStatus.textContent = data.error || '加载映射失败';
        }
      } catch (err) {
        refs.mappingStatus.textContent = '加载映射失败';
      }
    };

    refs.mappingToggle?.addEventListener('click', () => {
      const show = refs.mappingPanel.classList.contains('hidden');
      refs.mappingPanel.classList.toggle('hidden', !show);
      if (show && !mappingLoaded) {
        loadMapping();
        mappingLoaded = true;
      }
    });

    refs.mappingSave?.addEventListener('click', async () => {
      try {
        const parsed = JSON.parse(refs.mappingInput.value);
        const res = await fetch('/mapping', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ mappings: parsed })
        });
        const data = await res.json();
        if (data.ok) {
          refs.mappingStatus.textContent = '保存成功';
          log('映射配置已保存');
        } else {
          refs.mappingStatus.textContent = data.error || '保存失败';
        }
      } catch (err) {
        refs.mappingStatus.textContent = 'JSON 解析或保存失败';
      }
    });

    const cleanupAction = async (target, label) => {
      const formData = new FormData();
      formData.append('target', target);
      setStatus(`正在清理${label}...`);
      log(`开始清理${label}`);
      try {
        const res = await fetch('/cleanup', { method: 'POST', body: formData, credentials: 'same-origin' });
        const data = await res.json();
        if (!data.ok) {
          setStatus(data.error || `${label}清理失败`, true);
          log(`${label}清理失败`);
          return;
        }
        setStatus(`${label}清理完成，删除 ${data.removed} 项`);
        log(`${label}清理完成，删除 ${data.removed} 项`);
        if (data.errors && data.errors.length) {
          log(`清理错误: ${data.errors.join('; ')}`);
        }
      } catch (err) {
        setStatus(`${label}清理请求失败`, true);
        log(`${label}清理请求失败`);
      }
    };

    refs.cleanupLogs?.addEventListener('click', () => cleanupAction('logs', '日志'));
    refs.cleanupTemp?.addEventListener('click', () => cleanupAction('temp', '临时目录'));

    refs.mergeBtn.addEventListener('click', async () => {
      const files = [...filesState];
      if (!files.length) {
        setStatus('请先选择文件。', true);
        return;
      }

      const formData = new FormData();
      files.forEach(f => formData.append('files', f.file, f.name));
      if (document.getElementById('normalize').checked) formData.append('normalize_columns', 'on');
      if (document.getElementById('fuzzy').checked) formData.append('enable_fuzzy', 'on');
      if (document.getElementById('dedup').checked) formData.append('remove_duplicates', 'on');
      if (document.getElementById('smart').checked) formData.append('smart_dedup', 'on');
      const dedupKeys = document.getElementById('dedup_keys').value.trim();
      formData.append('dedup_keys', dedupKeys);
      const excluded = collectExcludedColumns();
      formData.append('exclude_columns', excluded.join(','));
      const fmt = document.querySelector('input[name="output_format"]:checked')?.value || 'xlsx';
      formData.append('output_format', fmt);

      setStatus('正在合并，请稍候...');
      refs.downloadBox.style.display = 'none';
      refs.mergeBtn.disabled = true;
      log('开始合并');

      try {
        const res = await fetch('/merge', {
          method: 'POST',
          body: formData,
          credentials: 'same-origin'
        });
        const data = await res.json();
        refs.mergeBtn.disabled = false;
        if (!data.ok) {
          setStatus(data.error || '合并失败', true);
          log(`合并失败：${data.error || '未知错误'}`);
          return;
        }
        setStatus('合并成功，点击下载结果。');
        refs.downloadBox.innerHTML = `<a href="${data.download_url}" target="_blank">下载结果文件</a>`;
        refs.downloadBox.style.display = 'block';
        log('合并完成，可下载结果');
      } catch (err) {
        refs.mergeBtn.disabled = false;
        setStatus('请求失败，请稍后重试。', true);
        log('合并请求失败');
      }
    });

    renderFiles();
    log('main.js init 完成');
  });
})();
