# Smartsheet Tracker - Improvement Roadmap

> **Important Note**: The groups (NA, NF, NH, NM, NP, NT, NV) are **PRODUCT GROUPS**, not regional groups.

This document outlines potential improvements for the Smartsheet Change Tracking and Reporting System used by the Noctua Returns Department for Amazon Content Management (ACM).

---

## Table of Contents
1. [High Priority](#high-priority)
2. [Code Quality & Architecture](#code-quality--architecture)
3. [New Features](#new-features)
4. [Performance Optimizations](#performance-optimizations)
5. [Error Handling & Reliability](#error-handling--reliability)
6. [Testing](#testing)
7. [Configuration Management](#configuration-management)
8. [Reporting Enhancements](#reporting-enhancements)
9. [Monitoring & Observability](#monitoring--observability)

---

## High Priority

### 1. Centralized Configuration Module
**Current State**: Configuration (SHEET_IDs, user lists, product group lists) is duplicated across multiple files (`smartsheet_tracker.py`, `smartsheet_report.py`, `smartsheet_status_updater.py`).

**Improvement**: Create a single `config.py` module that all scripts import from.

```python
# config.py
PRODUCT_GROUPS = ["NA", "NF", "NH", "NM", "NP", "NT", "NV"]
USERS = ["DM", "EK", "HI", "JHU", "LK", "SM"]
SHEET_IDS = { ... }
PHASE_FIELDS = [ ... ]
```

**Benefits**:
- Single source of truth
- Easier to add new product groups or users
- Reduced risk of inconsistencies

### 2. Type Hints and Dataclasses
**Current State**: Functions use dictionaries for structured data without type hints.

**Improvement**: Add Python type hints and use dataclasses for structured data.

```python
from dataclasses import dataclass
from typing import Optional
from datetime import date

@dataclass
class ChangeRecord:
    timestamp: str
    group: str
    row_id: int
    phase: int
    date_field: str
    date: date
    user: str
    marketplace: str
```

**Benefits**:
- Better IDE support and autocomplete
- Catch bugs earlier
- Self-documenting code

### 3. Unit Tests
**Current State**: No automated tests exist.

**Improvement**: Add pytest-based unit tests for critical functions:
- `parse_date()` - Date parsing logic
- `normalize_date_for_comparison()` - Date normalization
- `calculate_weekly_stats()` - Statistics calculation
- Report generation components

---

## Code Quality & Architecture

### 4. Extract Shared Utilities
Create a `utils/` directory with reusable modules:
- `utils/date_utils.py` - Date parsing and normalization
- `utils/smartsheet_client.py` - Smartsheet API wrapper
- `utils/csv_utils.py` - CSV reading/writing helpers

### 5. Dependency Injection for Smartsheet Client
**Current State**: Smartsheet client is created directly in functions.

**Improvement**: Pass client as parameter for better testability.

```python
def track_changes(client: smartsheet.Smartsheet, state: dict) -> TrackingResult:
    ...
```

### 6. Replace Magic Numbers/Strings
**Current State**: Sheet IDs and column names are hardcoded.

**Improvement**: Use enums or constants with descriptive names.

```python
class ProductGroup(Enum):
    NA = "NA"
    NF = "NF"
    # ...

class Phase(Enum):
    KONTROLLE = 1
    BE = 2
    K2 = 3
    C = 4
    REOPEN = 5
```

### 7. Structured Logging
**Current State**: Using basic logging with print-style messages.

**Improvement**: Use structured logging with context.

```python
logger.info("change_detected", extra={
    "group": group,
    "row_id": row.id,
    "phase": phase_no,
    "user": user_val
})
```

---

## New Features

### 8. Email Notifications
Send automated email reports when:
- Weekly/monthly reports are generated
- Error threshold is exceeded
- Unusual activity patterns detected

### 9. Slack/Teams Integration
Push notifications to team channels:
- Daily summary of changes
- Alert on tracking errors
- Report availability notifications

### 10. Interactive Dashboard
Create a simple web dashboard (Flask/FastAPI) showing:
- Real-time tracking status
- Change history visualization
- User activity leaderboards

### 11. Historical Comparison Reports
Compare metrics across time periods:
- Week-over-week comparisons
- Month-over-month trends
- Year-to-date summaries

### 12. Product Group Deep Dive Reports
Generate specialized reports for individual product groups:
- Dedicated PDF per product group
- Product-specific metrics and trends
- Marketplace-specific breakdowns

### 13. User Performance Analytics
Track and report on individual user metrics:
- Average items processed per day
- Phase completion patterns
- Productivity trends over time

### 14. Data Export Functionality
Export data in multiple formats:
- Excel with pivot tables
- JSON for API consumption
- CSV with configurable columns

---

## Performance Optimizations

### 15. Batch API Calls
**Current State**: Each row is processed individually.

**Improvement**: Use Smartsheet's bulk operations where possible.

### 16. Parallel Sheet Processing
**Current State**: Sheets are processed sequentially.

**Improvement**: Use `concurrent.futures` to process multiple sheets in parallel.

```python
from concurrent.futures import ThreadPoolExecutor

with ThreadPoolExecutor(max_workers=3) as executor:
    results = executor.map(process_sheet, SHEET_IDS.items())
```

### 17. Incremental State Updates
**Current State**: Full state file is written on every save.

**Improvement**: Implement incremental updates for large state files.

### 18. Caching Layer
Add caching for:
- Smartsheet column mappings (rarely change)
- Frequently accessed report data
- User/group color assignments

---

## Error Handling & Reliability

### 19. Retry Logic with Exponential Backoff
**Current State**: API errors cause immediate failure.

**Improvement**: Implement retry logic for transient failures.

```python
from tenacity import retry, stop_after_attempt, wait_exponential

@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=2, max=10))
def fetch_sheet(client, sheet_id):
    return client.Sheets.get_sheet(sheet_id)
```

### 20. Graceful Degradation
Continue processing other sheets if one fails, then report partial results.

### 21. Data Validation
Add validation for:
- CSV integrity checks
- State file corruption detection
- Date range sanity checks

### 22. Transaction-like State Management
Implement atomic state updates to prevent data corruption:
- Write to temp file first
- Validate before replacing
- Keep backup of previous state

---

## Testing

### 23. Unit Test Suite
Test coverage for:
- Date parsing (all formats)
- Statistics calculation
- Report data aggregation
- State management

### 24. Integration Tests
Test actual Smartsheet API interactions:
- Read operations
- Write operations
- Error scenarios

### 25. Mock Smartsheet Server
Create a mock server for testing without API calls:
- Faster test execution
- No API quota usage
- Deterministic test data

### 26. Report Visual Regression Tests
Compare generated PDFs against baseline:
- Layout consistency
- Chart rendering
- Color accuracy

---

## Configuration Management

### 27. Environment-Based Configuration
Support multiple environments:
- Development (sandbox Smartsheet)
- Staging (test data)
- Production (live data)

### 28. YAML Configuration File
Move configuration to `config.yaml`:

```yaml
product_groups:
  NA:
    name: "NA Product Group"
    sheet_id: 6141179298008964
    total_products: 1779
    color: "#E63946"
  NF:
    name: "NF Product Group"
    sheet_id: 615755411312516
    # ...

users:
  DM:
    full_name: "User DM"
    color: "#223459"
  # ...
```

### 29. Runtime Configuration Updates
Allow updating configuration without code changes:
- Add/remove product groups
- Add/remove users
- Adjust phase definitions

---

## Reporting Enhancements

### 30. Report Templates
Allow customizable report templates:
- Different layouts for different audiences
- Executive summary vs. detailed reports
- Configurable chart types

### 31. Interactive PDF Elements
Add hyperlinks and bookmarks to PDFs:
- Table of contents with links
- Cross-references between sections
- External links to Smartsheet

### 32. Multi-Language Support
Support report generation in multiple languages:
- German
- English
- Others as needed

### 33. Report Scheduling Flexibility
More granular scheduling options:
- Custom day of week for weekly reports
- Specific time zones
- On-demand generation with date ranges

### 34. Trend Analysis Charts
Add trend visualizations:
- Line charts for changes over time
- Moving averages
- Forecast projections (simple)

---

## Monitoring & Observability

### 35. Health Check Endpoint
Create a simple health check for monitoring:
- Last successful run
- Current state file status
- API connectivity check

### 36. Metrics Collection
Collect operational metrics:
- Execution duration
- Changes per run
- API call counts
- Error rates

### 37. Alerting Rules
Define alerting conditions:
- No changes detected for X days
- Error rate exceeds threshold
- Unusual activity patterns

### 38. Audit Trail Enhancement
Enhanced change tracking:
- Who initiated each run
- Configuration changes
- Manual corrections

---

## Implementation Priority Matrix

| Priority | Effort | Items |
|----------|--------|-------|
| High | Low | #1, #6, #20, #21 |
| High | Medium | #2, #3, #19, #28 |
| High | High | #4, #5, #23 |
| Medium | Low | #7, #33 |
| Medium | Medium | #8, #9, #14, #35 |
| Medium | High | #10, #11, #15, #16 |
| Low | Low | #31, #37 |
| Low | Medium | #12, #13, #34 |
| Low | High | #24, #25, #26, #32 |

---

## Quick Wins (Can be done immediately)

1. **Create `config.py`** - Extract all shared constants
2. **Add type hints** - Start with core functions
3. **Add basic unit tests** - Cover date parsing first
4. **Implement retry logic** - Wrap API calls with tenacity
5. **Add health check script** - Simple status verification

---

## Notes

- All improvements should maintain backwards compatibility with existing data
- Changes to the state file format should include migration scripts
- New features should be feature-flagged when possible
- Documentation should be updated alongside code changes

---

*Last Updated: 2026-01-19*
