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
 * Generate simple error ID
 */
function generateErrorId() {
  return `error_${Date.now()}`;
}

/**
 * Simple error logger
 * @param {string} message - Error message
 * @param {Object} context - Additional context
 * @returns {string} Error ID
 */
function logError(message, context = {}) {
  const errorId = generateErrorId();
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
 * Log warning
 * @param {string} message - Warning message
 * @param {Object} context - Additional context
 */
function logWarning(message, context = {}) {
  const timestamp = new Date().toISOString();

  const logEntry = {
    timestamp,
    message,
    ...context
  };

  console.warn('Warning:', JSON.stringify(logEntry));
}

/**
 * Log info
 * @param {string} message - Info message
 * @param {Object} context - Additional context
 */
function logInfo(message, context = {}) {
  const timestamp = new Date().toISOString();

  const logEntry = {
    timestamp,
    message,
    ...context
  };

  console.log('Info:', JSON.stringify(logEntry));
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

/**
 * Validate required parameters
 * @param {Object} params - Parameters to validate
 * @param {Array} required - Required parameter names
 * @returns {Object} Validation result
 */
function validateRequired(params, required) {
  const missing = required.filter(key => !params[key]);

  if (missing.length > 0) {
    return {
      isValid: false,
      message: `Missing required parameters: ${missing.join(', ')}`
    };
  }

  return { isValid: true };
}

// Error handling utilities for Zero-Dependency Architecture
// Individual functions exported directly for V8 optimization