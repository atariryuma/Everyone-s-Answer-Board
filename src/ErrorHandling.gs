function handleError(prefix, error, returnObject) {
  console.error(prefix + ':', error);
  var message = (error && error.message) ? error.message : String(error);
  if (returnObject) {
    return { status: 'error', message: 'エラーが発生しました: ' + message };
  }
  throw new Error(prefix + ': ' + message);
}

if (typeof module !== 'undefined') {
  module.exports = { handleError };
}
