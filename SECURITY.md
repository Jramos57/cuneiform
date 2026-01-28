# Security Policy

## Supported Versions

The following versions of Cuneiform are currently supported with security updates:

| Version | Supported          |
| ------- | ------------------ |
| 0.1.x   | :white_check_mark: |

## Reporting a Vulnerability

If you discover a security vulnerability in Cuneiform, please report it responsibly:

### Private Disclosure

**Do not file public issues for security vulnerabilities.** Instead, please use one of the following methods:

1. **GitHub Security Advisories** (preferred): 
   - Navigate to the [Security tab](https://github.com/jramos57/cuneiform/security)
   - Click "Report a vulnerability"
   - Provide detailed information about the vulnerability

2. **Email**: Send details to [jramos57] via GitHub
   - Include "SECURITY" in the subject line
   - Provide a detailed description of the vulnerability
   - Include steps to reproduce if possible
   - Mention any potential impact

### What to Include

When reporting a vulnerability, please include:

- **Description**: A clear description of the vulnerability
- **Impact**: What an attacker could potentially do
- **Reproduction Steps**: Detailed steps to reproduce the issue
- **Affected Versions**: Which versions of Cuneiform are affected
- **Suggested Fix** (optional): If you have ideas for how to fix it
- **Environment**: Swift version, platform, etc.

### Response Timeline

- **Initial Response**: Within 48 hours
- **Status Update**: Within 1 week
- **Fix Timeline**: Depends on severity and complexity

### Disclosure Policy

- Security vulnerabilities will be addressed promptly
- A fix will be released as soon as possible
- Credit will be given to reporters (unless anonymity is requested)
- Public disclosure will occur after a fix is available

## Security Considerations for Users

When using Cuneiform to process .xlsx files:

### File Input Validation

- Always validate file sources before processing
- Be cautious when opening .xlsx files from untrusted sources
- Consider sandboxing file processing operations

### Formula Evaluation

- The formula evaluator processes and executes formulas from .xlsx files
- While Cuneiform doesn't execute arbitrary code, complex formulas may consume significant CPU/memory
- Consider setting resource limits when processing untrusted files

### ZIP File Processing

- .xlsx files are ZIP archives; be aware of potential ZIP-related vulnerabilities
- Cuneiform uses Swift's native APIs for ZIP processing
- Very large or deeply nested ZIP structures may impact performance

### Dependencies

- Cuneiform has minimal dependencies (only swift-testing for development)
- Review the [Package.swift](Package.swift) file for the complete dependency list

## Known Limitations

Current known limitations that are not security vulnerabilities:

- Formula evaluation is limited to the 467 implemented functions
- No macros or VBA support (by design)
- No external data connections or web service functions

---

Thank you for helping keep Cuneiform secure!
