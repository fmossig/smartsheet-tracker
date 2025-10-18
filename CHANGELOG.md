# Changelog

All notable changes to the Smartsheet Change Tracking & Reporting System will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

## [Unreleased]

### Added
- Comprehensive documentation suite
  - README.md with full system overview
  - QUICKSTART.md for new users
  - ARCHITECTURE.md with technical details
  - EXAMPLES.md with practical use cases
  - CONTRIBUTING.md for contributors
  - This CHANGELOG.md
- System validation script (`validate_system.py`)
- Automated checks for dependencies, environment, data integrity

### Changed
- Improved .gitignore to properly track tracker_logs while excluding temporary files

## [2025-07-23] - Production System

### Core Features
The production system includes:

#### Data Collection
- **Incremental tracking** (`smartsheet_date_change_tracker.py`)
  - Daily tracking of date field changes
  - Monthly log rotation for manageable file sizes
  - Deduplication via backup state file
  - Bootstrap mode for initial data loading
  
- **Weekly snapshots** (`smartsheet_weekly_change_tracker.py`)
  - Last 7 days overview
  - Overwrites previous file
  
- **Status snapshots** (`smartsheet_status_snapshot.py`)
  - 30/60/90 day breakdowns
  - Summary statistics

#### Report Generation
- **Main status report** (`smartsheet_status_report.py`)
  - Professional PDF with ReportLab
  - Cover page with metadata
  - Product group overview with bar charts
  - Phase distribution pie charts
  - Per-group detailed pages with:
    - Employee contribution stacked bars
    - KPI boxes (article counts, percentages)
    - Country/marketplace rankings
  - Customizable time periods via environment variables
  
- **Periodic reports** (`smartsheet_periodic_status_report.py`)
  - Alternative report from tracker logs only
  - Simpler visualizations
  - Command-line driven (--week, --month, --from/--to)

#### Orchestration
- **Master orchestrator** (`smartsheet_reports_orchestrator.py`)
  - Unified interface for all modes
  - Modes: nightly, weekly, monthly, bootstrap
  - Automatic period calculation
  - Log trimming (keeps last N months)

#### Automation
- **GitHub Actions workflows**
  - Nightly tracking (00:00 UTC daily)
  - Weekly reports (Mondays 01:15 UTC)
  - Monthly reports (1st of month 02:00 UTC)
  - Bootstrap on demand

### Data Schema

#### Tracker Logs
```csv
Änderung am,Produktgruppe,Land/Marketplace,Phase,Feld,Datum,Mitarbeiter
```

Fields tracked:
- 7 product groups: NA, NF, NH, NM, NP, NT, NV
- 5 phases: Kontrolle, BE am, K am, C am, Reopen C2 am
- 6 employees: DM, EK, HI, JHU, LK, SM
- 17+ marketplaces: de, com, co.uk, fr, it, es, etc.

#### Storage Structure
```
tracker_logs/
  ├── date_changes_log_MM.YYYY.csv  (monthly rotation)
  └── date_backup.csv               (deduplication state)

reports/
  ├── weekly/YYYY/YYYY-Wxx/
  └── monthly/YYYY/

weekly/
  └── weekly_changes.csv

status/
  ├── status_report_YYYY-MM-DD.pdf
  ├── status_snapshot_YYYY-MM-DD.csv
  └── status_summary_YYYY-MM-DD.csv
```

### Configuration

#### Sheet IDs
- NA: 6141179298008964
- NF: 615755411312516
- NH: 123340632051588
- NP: 3009924800925572
- NT: 2199739350077316
- NV: 8955413669040004
- NM: 4275419734822788
- Seed: 6879355327172484

#### Color Schemes
Product groups, phases, and employees each have assigned colors for consistent visualization across all reports.

## Version History

### Version Format
Versions are dated: `YYYY-MM-DD_tracker`

Current version constant in code:
```python
CODE_VERSION = "2025-07-23_tracker"
```

### Key Milestones

- **2025-07-23**: Production system established
  - All core features operational
  - GitHub Actions automation live
  - Processing ~3,700+ tracked changes
  - Generating 5+ reports regularly

- **2025-10-18**: Documentation release
  - Comprehensive docs added
  - Validation script created
  - Contributing guidelines established
  - Examples and quickstart guides

## Migration Notes

### From Manual Tracking to Automated
If migrating from manual processes:

1. **Bootstrap the system**:
   ```bash
   python smartsheet_reports_orchestrator.py bootstrap --months 3
   ```

2. **Configure GitHub Actions**:
   - Add `SMARTSHEET_TOKEN` secret
   - Enable workflows

3. **Verify automation**:
   - Check nightly runs commit logs
   - Confirm weekly/monthly reports appear
   - Review artifacts in Actions tab

### Updating Product Groups
When adding new groups:

1. Add to `SHEET_IDS` in all scripts
2. Add color to `GROUP_COLORS`
3. Restart tracking to collect new data
4. Regenerate reports to include new group

### Schema Changes
If CSV format needs modification:

1. Create migration script
2. Test on copy of data
3. Update all parser functions
4. Document in this changelog
5. Provide rollback procedure

## Known Issues

### Current Limitations
- No automated tests (manual validation only)
- Limited error recovery in tracker
- Report generation can be slow for large datasets
- No web interface (PDF only)

### Workarounds
- **Slow generation**: Run reports during off-hours
- **Missing data**: Re-run bootstrap to fill gaps
- **Corrupted backup**: Delete and regenerate from logs

## Planned Features

### Under Consideration
- [ ] Web dashboard for live viewing
- [ ] Email delivery of reports
- [ ] Slack/Teams integration
- [ ] Alert system for anomalies
- [ ] Excel export functionality
- [ ] Interactive drill-down reports
- [ ] ML-based predictions
- [ ] Mobile-friendly viewer
- [ ] REST API for programmatic access
- [ ] Automated testing suite

### Requested Enhancements
- Custom report templates
- More granular time periods
- Additional chart types
- Comparison views (month-over-month)
- Employee performance metrics
- Project timeline visualizations

## Breaking Changes

### None Yet
This is the baseline version. Future breaking changes will be documented here with:
- Description of change
- Migration path
- Deprecation timeline
- Rollback procedure

## Security

### Advisories
No security issues reported yet.

### Best Practices
- Always use environment variables for tokens
- Never commit `.env` files
- Regularly rotate API tokens
- Review repository access permissions
- Monitor workflow logs for anomalies

## Support

### Getting Help
1. Check documentation (README, EXAMPLES, QUICKSTART)
2. Run validation: `python validate_system.py`
3. Search existing issues
4. Contact maintainers

### Reporting Issues
Include:
- Error messages/logs
- Steps to reproduce
- Expected vs actual behavior
- Environment details (Python version, OS)
- Related data (if safe to share)

## Credits

### Contributors
- System design and implementation: [Original team]
- Documentation: [Documentation contributors]
- Testing and feedback: [Test team]

### Dependencies
- [Smartsheet Python SDK](https://github.com/smartsheet-platform/smartsheet-python-sdk)
- [ReportLab](https://www.reportlab.com/)
- [python-dotenv](https://github.com/theskumar/python-dotenv)

### Acknowledgments
Special thanks to:
- Amazon Content Management team
- Noctua team
- All contributors and testers

---

## How to Update This Changelog

When making changes:

1. Add entry under `[Unreleased]` section
2. Use categories: Added, Changed, Deprecated, Removed, Fixed, Security
3. Write clear, user-focused descriptions
4. Include code examples for breaking changes
5. Reference issue numbers when applicable

When releasing:

1. Change `[Unreleased]` to `[YYYY-MM-DD]`
2. Update `CODE_VERSION` in source code
3. Create git tag
4. Add new `[Unreleased]` section at top

---

*For more information, see [CONTRIBUTING.md](CONTRIBUTING.md)*
