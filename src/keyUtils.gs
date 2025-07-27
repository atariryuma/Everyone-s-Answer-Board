/**
 * ユーザーIDと識別子から一意のキャッシュキーを生成します。
 * @param {string} prefix - キーのプレフィックス
 * @param {string} userId - 対象ユーザーID
 * @param {string} [suffix] - 追加識別子
 * @return {string} 生成されたキャッシュキー
 */
function buildUserScopedKey(prefix, userId, suffix) {
  if (!userId) throw new Error('userId is required for cache key');
  let key = `${prefix}_${userId}`;
  if (suffix) key += `_${suffix}`;
  return key;
}

