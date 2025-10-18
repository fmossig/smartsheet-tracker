# Contributing Guide

Thank you for considering contributing to the Smartsheet Change Tracking & Reporting System!

## Getting Started

1. **Fork the repository** (if external contributor)
2. **Clone your fork**:
   ```bash
   git clone https://github.com/your-username/smartsheet-tracker.git
   cd smartsheet-tracker
   ```

3. **Set up development environment**:
   ```bash
   # Install dependencies
   pip install -r requirements.txt
   
   # Configure environment
   cp .env.example .env  # Then edit with your token
   
   # Validate setup
   python validate_system.py
   ```

4. **Create a feature branch**:
   ```bash
   git checkout -b feature/your-feature-name
   ```

## Development Workflow

### Making Changes

1. **Keep changes focused**: One feature or fix per branch
2. **Follow existing code style**: Match the patterns already in use
3. **Test your changes**: Ensure they work before committing
4. **Update documentation**: If behavior changes, update relevant docs

### Code Style

#### Python
- Follow PEP 8 guidelines
- Use descriptive variable names
- Add docstrings for functions/classes
- Keep functions focused and small

**Example:**
```python
def parse_date_fuzzy(s):
    """
    Parse various date formats into datetime.date.
    
    Args:
        s (str): Date string to parse
        
    Returns:
        datetime.date or None: Parsed date or None if unparseable
    """
    if not s:
        return None
    # ... implementation
```

#### CSV Format
- Always preserve the standard column order
- Use UTF-8 encoding
- Include headers in all CSV files

#### Report Styling
- Use consistent colors (defined in dictionaries)
- Maintain mm units for sizing
- Keep professional appearance

### Testing

#### Manual Testing
```bash
# Test tracker
python smartsheet_date_change_tracker.py

# Test report generation
python smartsheet_status_report.py

# Test orchestrator
python smartsheet_reports_orchestrator.py nightly
```

#### Validation
```bash
# Run validation before committing
python validate_system.py

# Check Python syntax
python -m py_compile *.py
```

#### Data Integrity
```bash
# Verify CSV structure
head -1 tracker_logs/date_changes_log_*.csv

# Check for data corruption
python << 'EOF'
import csv, glob
for path in glob.glob("tracker_logs/*.csv"):
    with open(path) as f:
        try:
            list(csv.DictReader(f))
            print(f"✅ {path}")
        except Exception as e:
            print(f"❌ {path}: {e}")
EOF
```

### Committing

#### Commit Messages
Follow conventional commit format:

```
<type>(<scope>): <subject>

<body>

<footer>
```

**Types:**
- `feat`: New feature
- `fix`: Bug fix
- `docs`: Documentation changes
- `style`: Code style changes (formatting, etc.)
- `refactor`: Code refactoring
- `test`: Adding tests
- `chore`: Maintenance tasks

**Examples:**
```
feat(tracker): Add support for new phase field

Added tracking for new "Review" phase field that was added
to all product group sheets.

Closes #123

---

fix(report): Correct country ranking calculation

The ranking was using sum instead of average age. Fixed to
use average as intended.

---

docs(readme): Add section on custom report generation

Added examples showing how to generate reports for custom
date ranges and specific product groups.
```

#### Pre-commit Checklist
- [ ] Code is syntactically valid (`python -m py_compile *.py`)
- [ ] Changes are tested manually
- [ ] Documentation is updated if needed
- [ ] Commit message follows conventions
- [ ] No secrets or credentials in code
- [ ] No unnecessary files added (check `.gitignore`)

### Submitting Changes

1. **Push to your branch**:
   ```bash
   git push origin feature/your-feature-name
   ```

2. **Create Pull Request**:
   - Provide clear title and description
   - Reference any related issues
   - List what was changed and why
   - Include screenshots for UI changes

3. **Address review feedback**:
   - Respond to comments
   - Make requested changes
   - Push updates to same branch

## Types of Contributions

### Bug Fixes
- Fix data parsing errors
- Correct report generation issues
- Resolve automation failures
- Fix documentation errors

### Features
- New report visualizations
- Additional data tracking
- Enhanced aggregations
- Integration with other systems

### Documentation
- Improve README clarity
- Add usage examples
- Document edge cases
- Create troubleshooting guides

### Performance
- Optimize data processing
- Reduce API calls
- Speed up report generation
- Improve memory usage

### Automation
- New workflow triggers
- Enhanced error handling
- Better artifact management
- Notification systems

## Specific Guidelines

### Adding Product Groups

1. **Update sheet IDs**:
   ```python
   SHEET_IDS = {
       # ... existing ...
       "NX": 1234567890123456,  # New group
   }
   ```

2. **Add color**:
   ```python
   GROUP_COLORS = {
       # ... existing ...
       "NX": "#ABC123",  # Choose unique color
   }
   ```

3. **Update documentation**:
   - Add to README product group list
   - Update ARCHITECTURE if significant

4. **Test**:
   - Run tracker to collect data
   - Generate report to verify visualization

### Adding Phase Fields

1. **Update field definitions**:
   ```python
   PHASE_FIELDS = [
       # ... existing ...
       ("New Field", "New Field By", 6),
   ]
   ```

2. **Update schema documentation**:
   - Document new phase in README
   - Update ARCHITECTURE data flow

3. **Consider backward compatibility**:
   - Old logs won't have this field
   - Reports should handle missing data

### Modifying Report Layout

1. **Test with real data**:
   - Use actual tracker logs
   - Check with different data volumes
   - Verify on various time ranges

2. **Maintain consistency**:
   - Follow existing style patterns
   - Use defined color schemes
   - Keep sizing in mm units

3. **Consider performance**:
   - Large datasets take longer
   - Complex layouts increase generation time
   - Test with maximum expected data

### Changing CSV Format

⚠️ **Critical**: CSV format changes require careful handling

1. **Don't break existing data**:
   - Old logs must remain readable
   - Provide migration script if needed

2. **Update all readers**:
   - Tracker log loader functions
   - Report aggregation code
   - Validation scripts

3. **Document changes**:
   - Update schema in README
   - Note in CHANGELOG
   - Migration instructions

4. **Test thoroughly**:
   - Verify old data still works
   - Check new format is correct
   - Validate against all time periods

## Review Process

### What Reviewers Look For

1. **Correctness**: Does it work as intended?
2. **Quality**: Is code clean and maintainable?
3. **Testing**: Are changes verified?
4. **Documentation**: Are docs updated?
5. **Impact**: Any unintended side effects?

### Timeline

- Initial review: Within 3 business days
- Follow-up: Within 1 business day
- Approval: After all concerns addressed

## Release Process

### Versioning

Version format: `YYYY-MM-DD_tracker`

Update `CODE_VERSION` in `smartsheet_status_report.py`:
```python
CODE_VERSION = "2025-10-18_tracker"
```

### Changelog

Update CHANGELOG.md with:
- New features
- Bug fixes
- Breaking changes
- Migration notes

### Deployment

1. Merge to main branch
2. Tag release: `git tag v2025-10-18`
3. Push tag: `git push --tags`
4. GitHub Actions auto-deploys

## Best Practices

### Data Safety

- **Never delete data files** in production
- **Always backup** before major changes
- **Test destructive operations** on copies
- **Use version control** for all changes

### API Usage

- **Minimize calls**: Use incremental tracking
- **Handle rate limits**: Add retry logic
- **Cache when possible**: Store metadata
- **Graceful failures**: Don't lose data on errors

### Performance

- **Profile before optimizing**: Measure actual bottlenecks
- **Cache expensive operations**: Especially PDF generation
- **Use appropriate data structures**: Dict for lookups, list for iteration
- **Consider memory**: Large datasets need streaming

### Security

- **Never commit secrets**: Use environment variables
- **Validate input**: Especially from external sources
- **Sanitize output**: Prevent injection in reports
- **Review dependencies**: Check for vulnerabilities

### Documentation

- **Keep up to date**: Update with code changes
- **Be specific**: Include examples and edge cases
- **Consider audience**: Write for various skill levels
- **Check accuracy**: Test documented procedures

## Getting Help

### Questions

- Check [README.md](README.md) first
- Review [EXAMPLES.md](EXAMPLES.md) for similar use cases
- Search existing issues
- Ask in discussions (if enabled)

### Issues

When reporting bugs:
1. Describe expected vs actual behavior
2. Provide steps to reproduce
3. Include error messages/logs
4. Note your environment (Python version, OS, etc.)

### Discussions

For general questions, ideas, or feedback:
- Use GitHub Discussions (if available)
- Email the maintainers
- Join team chat (if applicable)

## Code of Conduct

### Our Standards

- **Be respectful**: Treat all contributors with respect
- **Be constructive**: Provide helpful feedback
- **Be collaborative**: Work together toward solutions
- **Be professional**: Maintain appropriate communication

### Unacceptable Behavior

- Harassment or discrimination
- Trolling or insulting comments
- Personal attacks
- Publishing others' private information
- Other unprofessional conduct

### Enforcement

Violations may result in:
1. Warning
2. Temporary ban
3. Permanent ban

Report issues to: [maintainer email]

## Recognition

Contributors will be:
- Listed in CONTRIBUTORS.md
- Mentioned in release notes
- Credited in significant features

Thank you for contributing! 🎉
