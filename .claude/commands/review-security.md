# Review Security Command

## Description
Comprehensive security audit and vulnerability assessment for the application.

## Usage
```
/review-security [scope]
```

Options:
- `auth` - Authentication and authorization review
- `input` - Input validation and sanitization review
- `data` - Data protection and privacy review
- `all` - Complete security audit (default)

## Implementation

This command performs security analysis across multiple layers:

### 1. Authentication & Authorization Security
- Session management validation
- OAuth2 flow security assessment
- Service Account token protection
- Access control matrix verification

### 2. Input Validation & Sanitization
- XSS prevention testing
- SQL injection protection (if applicable)
- File upload security (if applicable)
- Input length and format validation

### 3. Data Protection & Privacy
- Sensitive data handling review
- Encryption at rest validation
- Transmission security (HTTPS)
- Data retention policy compliance

### 4. Application Security
- Error message information leakage
- Debug information exposure
- Rate limiting implementation
- CSRF protection validation

### 5. Infrastructure Security
- Google Apps Script security best practices
- API endpoint protection
- Cache security validation
- Logging and monitoring security

## Security Checklist

### Critical Security Requirements
- [ ] All user inputs are validated and sanitized
- [ ] No sensitive data in logs or error messages
- [ ] Proper access control on all endpoints
- [ ] XSS protection implemented
- [ ] CSRF protection where applicable
- [ ] Rate limiting on critical operations
- [ ] Secure session management
- [ ] Principle of least privilege enforced

### Security Scoring

| Category | Weight | Score | Status |
|----------|--------|-------|--------|
| Authentication | 25% | 95/100 | ✅ |
| Input Validation | 20% | 92/100 | ✅ |
| Data Protection | 20% | 88/100 | ⚠️ |
| Access Control | 20% | 94/100 | ✅ |
| Infrastructure | 15% | 90/100 | ✅ |

## Expected Output

```
🔒 Security Review Results
├── 🔐 Authentication Security: ✅ PASS
├── 🛡️ Input Validation: ✅ PASS  
├── 📊 Data Protection: ⚠️ REVIEW NEEDED
├── 🚫 Access Control: ✅ PASS
├── 🏗️ Infrastructure: ✅ PASS
└── 📈 Overall Security Score: 92/100
```

## Integration

This command integrates with:
- SecurityService diagnostic functions
- Input validation testing
- Access control verification
- Error handling security checks