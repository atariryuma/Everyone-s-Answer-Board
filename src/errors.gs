/**
 * errors.gs - Simplified Error Handling
 *
 * 🎯 Responsibilities:
 * - Simple error logging
 * - Basic error ID generation
 * - Console logging with context
 *
 * Following Google Apps Script Best Practices:
 * - Simple, direct error handling
 * - Minimal overhead
 * - Clear logging
 */

/**
 * Simple error logger
 * @param {string} message - Error message
 * @param {Object} context - Additional context
 * @returns {string} Error ID
 */
function logError(message, context = {}) {
  const errorId = `error_${Date.now()}`;
  const timestamp = new Date().toISOString();

  const logEntry = {
    errorId,
    timestamp,
    message,
    ...context
  };

  console.error('Error:', JSON.stringify(logEntry));
  return errorId;
}



/**
 * Safe function executor with error handling
 * @param {Function} fn - Function to execute
 * @param {Object} context - Context for error logging
 * @returns {Object} Result with success flag
 */
function safeExecute(fn, context = {}) {
  try {
    const result = fn();
    return { success: true, data: result };
  } catch (error) {
    const errorId = logError(error.message, {
      ...context,
      stack: error.stack
    });
    return {
      success: false,
      error: error.message,
      errorId
    };
  }
}


// Error handling utilities for Zero-Dependency Architecture
// Individual functions exported directly for V8 optimization