# Project Brief: PowerShell User Data Backup/Restore Utility

## Overview
Rewrite of a PowerShell-based utility for comprehensive user data backup and restore operations, designed for IT administrators and support staff to manage user data migration and backup processes. The tool features a GUI interface and supports PowerShell version 3.0 while maintaining compatibility with version 1.

Original file used for examples is stored in documentation\UserBackupRefresh_Persist_1_original.ps1

New rewrite file is UserBackupRefresh_Persist_1.ps1

## Core Requirements

### Backup Process
- Automatically detect and backup user profile directories
- Support custom file/folder selection
- Capture network drive mappings and printer configurations
- Export browser favorites and Outlook signatures
- Handle special cases like OneNote and Sticky Notes
- Progress tracking with visual feedback

### Restore Process
- Restore user data to original or new locations
- Handle user profile directory restoration
- Reestablish mapped-network connections and printer mappings
- Automated configuration manager actions
- Certificate management for WiFi

## Key Functions

### Backup Core
- `Set-InitialStateAdminBackup`: Initialize backup environment
- `Start-Backup`: Execute backup operations
- `Set-InitialStateBackup`: Configure backup parameters

### Restore Core
- `Set-InitialStateRestore`: Initialize restore environment
- `Start-Restore`: Execute restore operations
- `Set-GPupdate`: Handle group policy updates

### Frontend
- `Initialize-Form`: Setup XAML-based GUI
- `Show-ProgressBar`: Display progress feedback
- `Update-ProgressBar`: Real-time progress updates
- `Show-UserPrompt`: User interaction handling

### Processing
- `Format-RegexSafe`: Path and string sanitization
- `Copy-File`: Enhanced file copy operations
- `Get-FilePath`: Path validation and processing

## Target Functionality

### Backup Operations
1. Quick access links capture
2. Outlook signatures management
3. Browser favorites export
   - Edge favorites export
   - Chrome favorites import to Edge
4. Sticky notes data backup
5. Network drive mapping preservation
6. Printer configuration backup

### Restore Operations
1. GPupdate execution in separate CMD instance
2. Configuration manager actions execution
3. User certificates for WiFi management
4. NVIDIA container service handling
5. Signature timestamp-based replacement

## Technical Scope
- PowerShell version compatibility (v1-v3)
- GUI implementation using XAML
- System integration points
  - Registry operations
  - File system access
  - Network resources
  - System services
  - Certificate management

## Development Guidelines
1. Error handling with user feedback
2. Comments on syntax and program logic
3. Progress tracking for all operations
4. Backup validation
5. Restore verification
6. User permission management
7. Cross-machine compatibility
