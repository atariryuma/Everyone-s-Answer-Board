/**
 * 簡易ログシステム - GAS ネイティブ console 使用
 * ULog から console ベースに移行済み
 */

// 後方互換性のためのパフォーマンス測定関数のみ保持
function createTimer(label) {
  const startTime = Date.now();
  return {
    label: label,
    startTime: startTime,
    end: () => {
      const duration = Date.now() - startTime;
      console.log(`Performance: ${label} - ${duration}ms`);
      return duration;
    }
  };
}

console.log('ログシステム初期化完了 - GASネイティブconsole使用');