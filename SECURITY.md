# Security Policy

## Reporting a Vulnerability

If you discover a security vulnerability in any of these scripts, please report it by:

1. **Opening a GitHub Issue** - For non-sensitive issues
2. **Direct Contact** - For sensitive vulnerabilities, please reach out directly through GitHub

## Security Considerations

These scripts are designed for system administration and often require elevated privileges:

- **Review before running** - Always review scripts before executing them in your environment
- **Test in isolation** - Test scripts in a non-production environment first
- **Understand the scope** - Some scripts make tenant-wide or system-wide changes
- **Credentials** - Never hardcode credentials; use secure credential management

## Best Practices

When using scripts from this repository:

1. Download from the official repository only
2. Verify script integrity before execution
3. Run with least-privilege principles where possible
4. Keep audit logs of script executions
5. Update to latest versions for security fixes
