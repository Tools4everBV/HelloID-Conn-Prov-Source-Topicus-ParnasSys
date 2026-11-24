# Change Log

All notable changes to this project will be documented in this file. The format is based on [Keep a Changelog](https://keepachangelog.com), and this project adheres to [Semantic Versioning](https://semver.org).

## [1.1.1] - 24-11-2026

### Fixed
- Student import failing when ParnasSys returns 0 students (empty XML nodes were returned as strings instead of XmlElements)
- Employee import adding loose hashtable entries instead of one person when ParnasSys returns exactly 1 employee
- Student import failing on group assignments or enrollments without a date
- Debug logging toggle referencing a non-existing `$c` variable

## [1.1.0] - 17-09-2024

### Added
- Aangemeld and SchoolType field
- Previous schoolyear


## [1.0.0] - 09-12-2022

Initial release
