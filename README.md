# PowerShell Scripts Collection

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://docs.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078D6.svg)](https://www.microsoft.com/windows)

A collection of useful PowerShell scripts for Windows and Microsoft 365 administration.

## Scripts

| Script | Description |
|--------|-------------|
| [Update-SensitivityLabel](./Update-SensitivityLabel/) | Bulk migrate Microsoft 365 sensitivity labels across SharePoint Online and OneDrive using Microsoft Graph metered API. Includes app registration (confidential client) support, discovery mode, dry-run, and a companion script for managing Site Collection Admin permissions. **Battle-tested in production.** |
| [Remove-Dell-Bloatware](./Remove-Dell-Bloatware/) | Comprehensive Dell bloatware removal script. Removes SupportAssist, Dell Optimizer, and other pre-installed Dell software. **Battle-tested on 1000+ machines.** |
| [Limit-PowerPointVersions](./Limit-PowerPointVersions/) | Limits version history for PowerPoint files across all SharePoint sites. Helps free up storage by trimming excessive version history while keeping recent versions. **Status: Untested** |

## Related Projects

| Project | Description |
|---------|-------------|
| [FileLabeler](https://github.com/OlaProeis/FileLabeler) | PowerShell GUI application for bulk applying sensitivity labels to local files. Features drag-and-drop, date preservation, and comprehensive reporting. |

## Usage

Each script folder contains its own README with detailed documentation, requirements, and examples.

## AI Disclaimer

Most of the code in this repository was written with AI assistance (GitHub Copilot, Claude, etc.). All scripts have been tested and validated in production environments, but please review and test in your own environment before deploying.

## Contributing

Feel free to submit issues or pull requests if you have improvements or new scripts to add.

## License

[MIT License](LICENSE) - feel free to use and modify these scripts as needed.
