# Security policy

Please do not disclose vulnerabilities or sensitive workbook content in a
public issue. Use GitHub's private vulnerability reporting feature for security
reports.

Workbook files are ZIP packages containing XML and binary parts. Treat
untrusted workbooks as untrusted input. This library preserves VBA and other
package parts but does not execute them.
