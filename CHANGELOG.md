# WinGet-Wrapper Changelog

## Overview
This document tracks all changes made to the ConfigForge Labs fork of WinGet-Wrapper since forking from the original [SorenLundt/WinGet-Wrapper](https://github.com/SorenLundt/WinGet-Wrapper) repository.

**Original Repository:** https://github.com/SorenLundt/WinGet-Wrapper  
**Fork Repository:** https://github.com/configforgelabs/WinGet-Wrapper  
**Fork Date:** 2024  
**Current Status:** Active Development

---

## Summary of Changes

### Major Enhancements

#### 1. **WPF GUI Implementation** (Latest Update)
- **Status:** ✅ Completed
- **Description:** Complete rewrite of the GUI from Windows Forms to WPF (Windows Presentation Foundation)
- **Benefits:**
  - Modern, cleaner user interface
  - Better performance and responsiveness
  - Improved styling capabilities
  - Enhanced user experience
  - Better support for future features

#### 2. **Launcher Refactoring**
- **Status:** ✅ Completed
- **Change:** Replaced `Start-GUI.exe` with `Start-GUI.cmd`
- **Rationale:** 
  - Removed dependency on compiled executable
  - Improved transparency and security (plain text script)
  - Easier to maintain and customize
  - Better for enterprise environments that scan binaries

#### 3. **Documentation Updates**
- **Status:** ✅ Completed
- **Changes:**
  - Updated README.md with new GUI screenshots
  - Improved formatting and clarity
  - Updated image references to reflect WPF GUI

#### 4. **App Registration Fixes**
- **Status:** ✅ Completed
- **Description:** Fixed issues with Azure App Registration process for InTune authentication
- **Impact:** Improved reliability when connecting to Microsoft Graph API

---

## Detailed Commit History

### Recent Commits (ConfigForge Fork)

| Commit | Description | Status |
|--------|-------------|--------|
| `157e17b` | General update | ✅ Complete |
| `6b199b3` | GUI Update with WPF | ✅ Complete |
| `c1204bf` | README update | ✅ Complete |
| `40cf55c` | Update image | ✅ Complete |
| `ba1ac47` | GUI Screenshot replaced | ✅ Complete |
| `648c5a0` | Replace exe with cmd | ✅ Complete |
| `b54cc66` | Formatting fix | ✅ Complete |
| `72bf8ec` | App Registration Fix | ✅ Complete |

### Upstream Sync Status
- **Last Sync:** Merged upstream changes including ARM64 support
- **Upstream Commit:** `9d078e4` - Remove arch from DesktopAppInstaller to support ARM64
- **Sync Status:** ✅ Up to date with upstream main features

---

## Breaking Changes

### None Currently
All changes maintain backward compatibility with the original WinGet-Wrapper functionality.

---

## Roadmap & Planned Features

### High Priority 🔴

#### 1. **Enhanced Error Handling**
- **Status:** 📋 Planned
- **Description:** Improve error messages and logging throughout the application
- **Priority:** High
- **Estimated Effort:** Medium
- **Benefits:**
  - Better troubleshooting capabilities
  - Clearer user feedback
  - Improved reliability

#### 2. **Multi-Tenant Support**
- **Status:** 📋 Planned
- **Description:** Allow managing multiple Intune tenants from single GUI
- **Priority:** High
- **Estimated Effort:** High
- **Benefits:**
  - MSP-friendly functionality
  - Simplified multi-tenant management
  - Credential management improvements

#### 3. **Package Validation & Testing**
- **Status:** 📋 Planned
- **Description:** Pre-deployment validation and testing framework
- **Priority:** High
- **Estimated Effort:** High
- **Features:**
  - Test WinGet package before InTune upload
  - Validate detection scripts
  - Check requirement scripts
  - Verify package metadata

### Medium Priority 🟡

#### 4. **Bulk Operations Enhancement**
- **Status:** 📋 Planned
- **Description:** Improve batch processing capabilities
- **Priority:** Medium
- **Estimated Effort:** Medium
- **Features:**
  - Bulk edit package settings
  - Mass version updates
  - Batch deployment to multiple groups

#### 5. **Package Templates & Presets**
- **Status:** 📋 Planned
- **Description:** Predefined configurations for common scenarios
- **Priority:** Medium
- **Estimated Effort:** Low
- **Features:**
  - Save/load package configurations
  - Share templates across team
  - Industry-standard presets

#### 6. **Advanced Search & Filtering**
- **Status:** 📋 Planned
- **Description:** Enhanced package discovery capabilities
- **Priority:** Medium
- **Estimated Effort:** Medium
- **Features:**
  - Filter by publisher, category, size
  - Search history
  - Favorite packages
  - Package recommendations

#### 7. **Reporting & Analytics**
- **Status:** 📋 Planned
- **Description:** Deployment statistics and insights
- **Priority:** Medium
- **Estimated Effort:** High
- **Features:**
  - Deployment success rates
  - Package usage statistics
  - Version distribution reports
  - Export to Excel/CSV

### Low Priority 🟢

#### 8. **Dark Mode Theme**
- **Status:** 📋 Planned
- **Description:** Dark theme option for WPF GUI
- **Priority:** Low
- **Estimated Effort:** Low
- **Benefits:**
  - User preference support
  - Reduced eye strain
  - Modern UI experience

#### 9. **Localization Support**
- **Status:** 📋 Planned
- **Description:** Multi-language support (starting with German, French, Spanish)
- **Priority:** Low
- **Estimated Effort:** High
- **Current Limitation:** WinGet requires en-US client language

#### 10. **API/CLI Interface**
- **Status:** 📋 Planned
- **Description:** Programmatic access to WinGet-Wrapper functionality
- **Priority:** Low
- **Estimated Effort:** High
- **Use Cases:**
  - CI/CD integration
  - Automation scripts
  - Third-party tool integration

---

## Technical Debt & Maintenance

### Current Technical Debt

#### 1. **Code Documentation**
- **Issue:** Limited inline documentation in PowerShell scripts
- **Impact:** Medium
- **Priority:** Medium
- **Estimated Effort:** Medium
- **Action:** Add comprehensive comments and help documentation

#### 2. **Test Coverage**
- **Issue:** No automated testing framework
- **Impact:** High
- **Priority:** High
- **Estimated Effort:** High
- **Action:** Implement Pester tests for PowerShell scripts

#### 3. **Configuration Management**
- **Issue:** Settings scattered across scripts
- **Impact:** Low
- **Priority:** Low
- **Estimated Effort:** Medium
- **Action:** Centralize configuration in config file

### Dependency Updates Needed

| Dependency | Current Version | Latest Version | Priority |
|------------|----------------|----------------|----------|
| IntuneWin32App | Unknown | Check | Medium |
| Microsoft.Graph.Intune | Unknown | Check | Medium |
| Microsoft.Graph.Authentication | Unknown | Check | Medium |

---

## Known Issues & Limitations

### From Original Repository

1. **Language Requirement**
   - **Issue:** Client language must be en-US
   - **Reason:** WinGet output parsing is English-only
   - **Workaround:** Set en-US as highest priority language
   - **Status:** Upstream limitation

2. **System Context Challenges**
   - **Issue:** WinGet.exe location changes with versions
   - **Solution:** Dynamic path detection implemented
   - **Status:** ✅ Solved

### ConfigForge Fork Specific

1. **WPF GUI .NET Framework Requirement**
   - **Issue:** Requires .NET Framework 4.7.2 or higher
   - **Impact:** May not work on older systems
   - **Workaround:** Update .NET Framework
   - **Status:** By design

---

## Contributing to This Fork

### Development Workflow

1. Fork from `configforgelabs/WinGet-Wrapper`
2. Create feature branch: `git checkout -b feature/your-feature`
3. Make changes and test thoroughly
4. Commit with clear messages: `git commit -m "Add: feature description"`
5. Push to your fork: `git push origin feature/your-feature`
6. Create Pull Request to `configforgelabs/WinGet-Wrapper`

### Testing Requirements

Before submitting PR, ensure:
- [ ] GUI launches without errors
- [ ] Package search functionality works
- [ ] InTune import completes successfully
- [ ] Detection scripts function correctly
- [ ] CSV import/export works
- [ ] No breaking changes to existing functionality

### Code Style Guidelines

- **PowerShell:** Follow [PowerShell Best Practices](https://docs.microsoft.com/en-us/powershell/scripting/developer/cmdlet/strongly-encouraged-development-guidelines)
- **XAML:** Keep WPF code organized and commented
- **Documentation:** Update README.md and CHANGELOG.md for significant changes

---

## Version History

### v2.0.0 (ConfigForge - WPF Update) - Current
- ✅ Complete WPF GUI rewrite
- ✅ Replaced exe launcher with cmd script
- ✅ Updated documentation
- ✅ Fixed App Registration issues
- ✅ Synced with upstream ARM64 support

### v1.x (Original SorenLundt)
- Original Windows Forms GUI
- Core WinGet-Wrapper functionality
- CSV import/export
- InTune integration
- Detection and requirement scripts

---

## Upstream Sync Strategy

### Current Status
- **Sync Frequency:** As needed, when significant upstream changes occur
- **Last Sync:** ARM64 support merge
- **Conflicts:** None currently

### Sync Process
```bash
# Fetch upstream changes
git fetch upstream

# Merge upstream into main
git checkout main
git merge upstream/main

# Resolve conflicts if any
# Test thoroughly
# Push to origin
git push origin main
```

### Upstream Features to Monitor
- New WinGet command-line options
- Changes to package metadata structure
- Microsoft Graph API updates
- InTune Win32 app format changes

---

## Migration Guide

### For Users of Original WinGet-Wrapper

#### Switching to ConfigForge Fork

1. **Backup Current Setup**
   ```powershell
   # Backup your CSV files and custom scripts
   Copy-Item .\WinGet-WrapperImportFromCSV.csv .\backup\
   ```

2. **Clone ConfigForge Fork**
   ```bash
   git clone https://github.com/configforgelabs/WinGet-Wrapper.git
   cd WinGet-Wrapper
   ```

3. **Launch New GUI**
   ```cmd
   Start-GUI.cmd
   ```

4. **Import Existing Packages**
   - Use CSV import feature with your backed-up CSV file
   - Or manually re-add packages through search

#### What Stays the Same
- ✅ All PowerShell scripts remain compatible
- ✅ Detection scripts work unchanged
- ✅ InTune integration same process
- ✅ CSV format unchanged

#### What's Different
- 🎨 New WPF GUI interface (more modern look)
- 🚀 Launcher changed from `.exe` to `.cmd`
- 📸 Updated screenshots in documentation

---

## Support & Resources

### Getting Help

- **Issues:** Report bugs at https://github.com/configforgelabs/WinGet-Wrapper/issues
- **Discussions:** Q&A and feature requests in GitHub Discussions
- **Documentation:** README.md and this CHANGELOG.md

### Related Resources

- [Original WinGet-Wrapper](https://github.com/SorenLundt/WinGet-Wrapper)
- [WinGet Documentation](https://docs.microsoft.com/en-us/windows/package-manager/winget/)
- [InTune Win32 App Management](https://docs.microsoft.com/en-us/mem/intune/apps/apps-win32-app-management)
- [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/overview)

---

## License

This fork maintains the same license as the original WinGet-Wrapper repository.  
See [LICENSE.txt](LICENSE.txt) for full details.

---

## Acknowledgments

- **Original Author:** [Soren Lundt](https://github.com/SorenLundt) for creating WinGet-Wrapper
- **Contributors:** All contributors to the original repository
- **ConfigForge Labs:** For maintaining and enhancing this fork

---

## Changelog Metadata

**Last Updated:** January 2025  
**Document Version:** 1.0  
**Maintained By:** ConfigForge Labs  
**Status:** Active Development
