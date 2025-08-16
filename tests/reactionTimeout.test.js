const StudyQuestApp = {};

// Minimal implementation of sendReactionToServer with timeout logic
StudyQuestApp.CONSTANTS = {REACTION_TIMEOUT_MS: 50};

async function sendReactionToServer(rowIndex, reaction) {
  const timeoutMs = StudyQuestApp.CONSTANTS.REACTION_TIMEOUT_MS;
  const timeoutPromise = new Promise((_, reject) =>
    setTimeout(
      () => reject(new Error('リアクション送信がタイムアウトしました')),
      timeoutMs,
    ),
  );
  const gas = {addReaction: () => new Promise(() => {})};
  return Promise.race([
    gas.addReaction(rowIndex, reaction, 'Sheet1'),
    timeoutPromise,
  ]);
}

describe('sendReactionToServer timeout', () => {
  test('rejects if GAS call does not resolve in time', async () => {
    await expect(sendReactionToServer(1, 'LIKE')).rejects.toThrow(
      'タイムアウト',
    );
  });
});
