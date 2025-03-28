# Knowledge Transfer: Backup Tool Improvements

## Current Development Focus
We are improving the backup tool's handling of file paths and names, focusing on:
- Preserving original filenames during backup operations
- Optimizing path loading and management
- Maintaining proper backup timing and initialization

## Key Changes Made

### Get-BackupPaths.ps1 Improvements
- Moved from friendly names to actual filenames (e.g., "plum.sqlite" instead of "Sticky Notes")
- Simplified path array structure
- Added automatic file/folder type detection
- Improved error handling for inaccessible paths

### Main Script Integration
1. Path Loading Timing
   - Loads paths at window initialization for backup mode
   - Preserves paths during backup folder creation
   - Better handling of restore mode path detection

2. File Processing
   - Maintains original file structure
   - Uses relative paths for proper directory recreation
   - Improved progress tracking during operations

3. Error Handling
   - Better warnings for inaccessible paths
   - More informative error messages
   - Graceful handling of missing files/folders

## Technical Context

### Why These Changes Were Needed
1. Original Issue:
   - Backup was using friendly names instead of actual filenames
   - This caused problems during restore operations
   - Made file tracking and verification difficult

2. Path Loading Problems:
   - Paths were being reset during backup folder creation
   - Timing issues with when paths were loaded
   - Inconsistent path handling between backup/restore modes

### Implementation Details
1. Path Handling:
   ```powershell
   # Old approach (using friendly names)
   @{Path = $path; Name = "Friendly Name"; Type = "Folder"}

   # New approach (using actual names)
   @{
       Name = Split-Path $path -Leaf  # Gets actual filename
       Path = $path
       Type = if (Test-Path -Path $path -PathType Container) { "Folder" } else { "File" }
   }
   ```

2. Initialization Order:
   - Load default paths at window creation
   - Initialize UI elements
   - Set up event handlers
   - Handle mode-specific operations

## Future Considerations

### Watch Points
1. Path Changes:
   - Monitor system paths for Windows updates
   - Check for app-specific path changes
   - Verify path accessibility

2. File Handling:
   - Large file performance
   - Network path timeouts
   - Permission issues

### Potential Improvements
1. Performance:
   - Parallel file processing
   - Progress estimation
   - Cancel operation support

2. Usability:
   - Better path validation
   - Custom path management
   - Backup verification

### Maintenance Notes
- Keep path definitions up to date
- Test with various file types
- Monitor Windows path changes
- Check for app updates affecting paths
