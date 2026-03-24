(() => {
  const ready = (fn) => {
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', fn);
    } else {
      fn();
    }
  };

  ready(() => {
    const modalHelpers = window.DownloadModalHelpers || {
      buildInitialState: () => ({
        confirmedFilename: '',
        canDownload: false,
        statusText: '请先输入文件名，再点击“确定”。',
      }),
      buildConfirmedState: (filename, extension) => ({
        confirmedFilename: filename,
        canDownload: true,
        statusText: `已确认文件名：${filename}${extension}`,
      }),
      buildEditedState: (inputValue) => ({
        confirmedFilename: '',
        canDownload: false,
        statusText: inputValue.trim() ? '文件名已变更，请重新点击“确定”。' : '文件名不能为空。',
      }),
    };

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
      downloadModal: document.getElementById('download-modal'),
      downloadFilenameInput: document.getElementById('download-filename'),
      downloadExtension: document.getElementById('download-extension'),
      downloadConfirmBtn: document.getElementById('download-confirm'),
      downloadStartBtn: document.getElementById('download-start'),
      downloadCancelBtn: document.getElementById('download-cancel'),
      downloadStatus: document.getElementById('download-status'),
    };

    if (!refs.fileInput || !refs.fileList || !refs.inspectBtn) {
      console.error('初始化失败，DOM 元素缺失');
      return;
    }

    let filesState = [];
    let fileId = 1;
    let lastPreviews = [];
    let mappingLoaded = false;
    let pendingDownload = null;
    let confirmedDownloadFilename = '';
    let mergePollTimer = null;

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

    const clearNode = (node) => {
      while (node.firstChild) {
        node.removeChild(node.firstChild);
      }
    };

    const getOutputFormat = () => {
      return document.querySelector('input[name="output_format"]:checked')?.value || 'xlsx';
    };

    const stopMergePolling = () => {
      if (mergePollTimer) {
        window.clearTimeout(mergePollTimer);
        mergePollTimer = null;
      }
    };

    const closeDownloadModal = () => {
      if (!refs.downloadModal) return;
      refs.downloadModal.classList.add('hidden');
      pendingDownload = null;
      confirmedDownloadFilename = '';
      if (refs.downloadStartBtn) refs.downloadStartBtn.disabled = true;
      if (refs.downloadStatus) refs.downloadStatus.textContent = '';
    };

    const openDownloadModal = (downloadUrl, suggestedFilename, format) => {
      if (!refs.downloadModal || !refs.downloadFilenameInput || !refs.downloadExtension) {
        window.open(downloadUrl, '_blank', 'noopener');
        return;
      }
      pendingDownload = { downloadUrl, format };
      confirmedDownloadFilename = '';
      refs.downloadFilenameInput.value = suggestedFilename || 'merged';
      refs.downloadExtension.textContent = `.${format}`;
      const initialState = modalHelpers.buildInitialState();
      if (refs.downloadStartBtn) refs.downloadStartBtn.disabled = !initialState.canDownload;
      if (refs.downloadStatus) refs.downloadStatus.textContent = initialState.statusText;
      refs.downloadModal.classList.remove('hidden');
      refs.downloadFilenameInput.focus();
      refs.downloadFilenameInput.select();
    };

    const confirmDownloadFilename = () => {
      if (!pendingDownload) return;
      const filename = refs.downloadFilenameInput?.value?.trim() || '';
      if (!filename) {
        if (refs.downloadStatus) refs.downloadStatus.textContent = '文件名不能为空。';
        refs.downloadFilenameInput?.focus();
        return;
      }
      const confirmedState = modalHelpers.buildConfirmedState(
        filename,
        refs.downloadExtension?.textContent || '',
      );
      confirmedDownloadFilename = confirmedState.confirmedFilename;
      if (refs.downloadStartBtn) refs.downloadStartBtn.disabled = !confirmedState.canDownload;
      if (refs.downloadStatus) {
        refs.downloadStatus.textContent = confirmedState.statusText;
      }
      log(`已确认下载文件名：${confirmedDownloadFilename}${refs.downloadExtension?.textContent || ''}`);
    };

    const triggerDownload = () => {
      if (!pendingDownload) return;
      const filename = confirmedDownloadFilename || refs.downloadFilenameInput?.value?.trim() || '';
      if (!filename) {
        if (refs.downloadStatus) refs.downloadStatus.textContent = '请先输入并确认文件名。';
        refs.downloadFilenameInput?.focus();
        return;
      }
      const url = new URL(pendingDownload.downloadUrl, window.location.origin);
      if (filename) {
        url.searchParams.set('filename', filename);
      }
      const link = document.createElement('a');
      link.href = url.toString();
      link.style.display = 'none';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      log(`开始下载：${filename || '默认文件名'}${refs.downloadExtension?.textContent || ''}`);
      closeDownloadModal();
    };

    const renderDownloadButton = (downloadUrl, suggestedFilename, format) => {
      clearNode(refs.downloadBox);
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'btn secondary';
      button.textContent = '下载结果文件';
      button.addEventListener('click', () => {
        openDownloadModal(
          downloadUrl,
          suggestedFilename || 'merged',
          format || getOutputFormat(),
        );
      });
      refs.downloadBox.appendChild(button);
      refs.downloadBox.style.display = 'block';
    };

    const pollMergeStatus = async (taskId, statusUrl, fallbackFormat) => {
      stopMergePolling();
      try {
        const res = await fetch(statusUrl, { credentials: 'same-origin' });
        const data = await res.json();
        if (!res.ok || !data.ok) {
          refs.mergeBtn.disabled = false;
          setStatus(data.error || `任务查询失败（${res.status}）`, true);
          log(`合并失败：${data.error || `状态查询 HTTP ${res.status}`}`);
          return;
        }

        if (data.status === 'queued' || data.status === 'running') {
          setStatus('后台正在合并，请稍候...');
          log(`任务处理中：${data.status}`);
          mergePollTimer = window.setTimeout(() => {
            pollMergeStatus(taskId, statusUrl, fallbackFormat);
          }, 2000);
          return;
        }

        refs.mergeBtn.disabled = false;
        if (data.status === 'completed') {
          setStatus('合并成功，点击下载结果。');
          renderDownloadButton(
            data.download_url,
            data.suggested_filename,
            data.format || fallbackFormat,
          );
          log('合并完成，可下载结果');
          return;
        }

        setStatus(data.error || '合并失败', true);
        log(`合并失败：${data.error || data.status || '未知错误'}`);
      } catch (err) {
        refs.mergeBtn.disabled = false;
        setStatus('合并状态查询失败，请稍后重试。', true);
        log('合并状态查询失败');
      }
    };

    const renderFiles = () => {
      if (!filesState.length) {
        refs.fileList.textContent = '尚未选择文件。';
        return;
      }
      clearNode(refs.fileList);
      filesState.forEach((f) => {
        const label = document.createElement('label');
        label.style.display = 'flex';
        label.style.gap = '8px';
        label.style.alignItems = 'center';
        label.style.marginBottom = '4px';

        const input = document.createElement('input');
        input.type = 'checkbox';
        input.className = 'file-item';
        input.dataset.id = String(f.id);

        const span = document.createElement('span');
        span.textContent = `${f.name} — ${(f.size/1024).toFixed(1)} KB`;

        label.appendChild(input);
        label.appendChild(span);
        refs.fileList.appendChild(label);
      });
      refs.fileList.querySelectorAll('.file-item').forEach(cb => {
        cb.addEventListener('change', () => showPreviewForSelection());
      });
    };

    const renderColumns = (columns) => {
      if (!columns || !columns.length) {
        refs.columnsBox.textContent = '未获取到列信息。';
        return;
      }
      clearNode(refs.columnsBox);
      columns.forEach((col) => {
        const label = document.createElement('label');
        label.style.display = 'flex';
        label.style.gap = '8px';
        label.style.alignItems = 'center';
        label.style.marginBottom = '6px';

        const input = document.createElement('input');
        input.type = 'checkbox';
        input.className = 'col-item';
        input.value = col.name;
        input.disabled = !!col.is_meta;

        const span = document.createElement('span');
        span.textContent = col.name;

        const small = document.createElement('small');
        small.style.color = 'var(--muted)';
        small.textContent = `(${(col.sources || []).join(', ')})${col.is_meta ? ' - 保留' : ''}`;

        label.appendChild(input);
        label.appendChild(span);
        label.appendChild(small);
        refs.columnsBox.appendChild(label);
      });
    };

    const renderPreview = (previews) => {
      lastPreviews = previews || [];
      if (!lastPreviews.length) {
        refs.previewArea.textContent = '未获取到预览数据。';
        return;
      }
      clearNode(refs.previewArea);
      lastPreviews.forEach((p) => {
        const rows = p.rows || [];
        const header = (p.columns || []).join(' | ');
        const body = rows.map(r => Object.values(r).join(' | ')).join('\n');

        const wrapper = document.createElement('div');
        wrapper.style.marginBottom = '12px';

        const title = document.createElement('strong');
        title.textContent = `${p.file} / ${p.sheet}`;

        const pre = document.createElement('pre');
        pre.style.margin = '6px 0';
        pre.style.whiteSpace = 'pre-wrap';
        pre.textContent = `${header}\n${body}`;

        wrapper.appendChild(title);
        wrapper.appendChild(pre);
        refs.previewArea.appendChild(wrapper);
      });
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
      closeDownloadModal();
      stopMergePolling();
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
        setStatus(`${label}清理完成，删除 ${data.removed} 项，跳过 ${data.skipped || 0} 项`);
        log(`${label}清理完成，删除 ${data.removed} 项，跳过 ${data.skipped || 0} 项`);
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
    refs.downloadCancelBtn?.addEventListener('click', closeDownloadModal);
    refs.downloadConfirmBtn?.addEventListener('click', confirmDownloadFilename);
    refs.downloadStartBtn?.addEventListener('click', triggerDownload);
    refs.downloadModal?.addEventListener('click', (event) => {
      if (event.target === refs.downloadModal) {
        closeDownloadModal();
      }
    });
    refs.downloadFilenameInput?.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        confirmDownloadFilename();
      } else if (event.key === 'Escape') {
        closeDownloadModal();
      }
    });
    refs.downloadFilenameInput?.addEventListener('input', (event) => {
      const nextValue = event.target?.value || '';
      if (!confirmedDownloadFilename && !nextValue.trim()) {
        return;
      }
      const editedState = modalHelpers.buildEditedState(nextValue);
      confirmedDownloadFilename = editedState.confirmedFilename;
      if (refs.downloadStartBtn) refs.downloadStartBtn.disabled = !editedState.canDownload;
      if (refs.downloadStatus) refs.downloadStatus.textContent = editedState.statusText;
    });

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
      const fmt = getOutputFormat();
      formData.append('output_format', fmt);

      setStatus('正在合并，请稍候...');
      refs.downloadBox.style.display = 'none';
      stopMergePolling();
      refs.mergeBtn.disabled = true;
      log('开始合并');

      try {
        const res = await fetch('/merge', {
          method: 'POST',
          body: formData,
          credentials: 'same-origin'
        });
        const data = await res.json();
        if (!data.ok) {
          refs.mergeBtn.disabled = false;
          setStatus(data.error || '合并失败', true);
          log(`合并失败：${data.error || '未知错误'}`);
          return;
        }
        setStatus('文件已上传，后台正在合并...');
        log(`合并任务已创建：${data.task_id}`);
        pollMergeStatus(data.task_id, data.status_url, fmt);
      } catch (err) {
        refs.mergeBtn.disabled = false;
        setStatus('请求失败，请稍后重试。', true);
        log('合并请求失败');
      }
    });

    renderFiles();
    refs.downloadBox.style.display = 'none';
    closeDownloadModal();
    log('main.js init 完成');
  });
})();
