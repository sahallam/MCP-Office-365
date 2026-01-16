# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.1.0] - 2026-01-16

### Added
- Automatic timezone conversion for calendar events
  - System timezone detection for macOS, Linux, and Windows
  - Calendar events now display in local time instead of UTC
  - Supports all calendar operations: list, get, create, update

### Fixed
- Calendar events now properly convert time values from UTC to local timezone (not just relabel)
- Example: 04:00 UTC correctly shows as 15:00 AEDT (Australia/Melbourne) instead of 04:00

## [1.0.0] - Previous Release

### Features
- Outlook/Email management
- Calendar integration
- OneDrive file operations
- SharePoint site management
- Microsoft Teams integration
- Excel workbook operations
- Word document handling
- OneNote notebook management
- App-only and delegated authentication modes
- Persistent token caching with encryption
