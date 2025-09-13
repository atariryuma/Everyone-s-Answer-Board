# Check Authentication Flow

Verify that the web application follows the correct authentication flow pattern.

## Usage
```
/check-flow
```

## What This Command Does

1. **Verify Entry Point Behavior**
   - `/exec` access should → Login Page
   - NOT directly to main board

2. **Test Authentication Chain**
   ```
   /exec → Login → Setup (if needed) → Main Board
   ```

3. **Check Critical Issues**
   - `userId: undefined` errors
   - Direct main board access without auth
   - Broken page transitions

## Implementation

```javascript
// Test the main.gs routing logic
function testAuthFlow() {
  console.log('Testing /exec access...');

  // Should go to login first
  const result = doGet({parameter: {}});

  // Verify login page is shown
  if (result.getContent().includes('LoginPage')) {
    console.log('✅ Correct: /exec → Login Page');
  } else {
    console.log('❌ Error: /exec bypassing login');
  }
}
```

## Expected Behavior
- New users: Login → Setup → Main Board
- Returning users: Login → Main Board
- Never: Direct access to Main Board