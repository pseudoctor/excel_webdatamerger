(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
    return;
  }
  root.DownloadModalHelpers = factory();
}(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  function buildInitialState() {
    return {
      confirmedFilename: '',
      canDownload: false,
      statusText: '请先输入文件名，再点击“确定”。',
    };
  }

  function buildConfirmedState(filename, extension) {
    return {
      confirmedFilename: filename,
      canDownload: true,
      statusText: '已确认文件名：' + filename + extension,
    };
  }

  function buildEditedState(inputValue) {
    if (!inputValue.trim()) {
      return {
        confirmedFilename: '',
        canDownload: false,
        statusText: '文件名不能为空。',
      };
    }
    return {
      confirmedFilename: '',
      canDownload: false,
      statusText: '文件名已变更，请重新点击“确定”。',
    };
  }

  return {
    buildInitialState: buildInitialState,
    buildConfirmedState: buildConfirmedState,
    buildEditedState: buildEditedState,
  };
}));
