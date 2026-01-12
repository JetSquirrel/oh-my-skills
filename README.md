# Oh My Skills

A collection of Claude AI skills for various tasks. This repository is inspired by [Anthropic's skills repository](https://github.com/anthropics/skills).

## What are Skills?

Skills are packaged instructions and code that help Claude AI perform specific tasks more effectively. Each skill contains:
- A `SKILL.md` file with metadata and instructions
- Optional scripts and resources
- Example usage and documentation

## Available Skills

### Excel Sheet Reference

Create Excel files with multiple sheets and cross-sheet formula references.

**Features:**
- Multi-sheet Excel workbooks
- Cross-sheet references (Sheet1!A1:A10)
- COUNTIFS, VLOOKUP, MATCH formulas
- INDEX-MATCH combinations
- Error handling with IFERROR

**Location:** [`skills/excel-sheet-reference/`](./skills/excel-sheet-reference/)

**Quick Start:**
```bash
pip install openpyxl
python skills/excel-sheet-reference/scripts/create_excel_with_references.py
```

## Repository Structure

```
oh-my-skills/
├── skills/                    # All skills directory
│   └── excel-sheet-reference/ # Excel cross-sheet reference skill
│       ├── SKILL.md          # Skill definition and instructions
│       ├── EXAMPLES.md       # Additional examples and use cases
│       └── scripts/          # Python scripts for Excel manipulation
│           └── create_excel_with_references.py
├── LICENSE                    # MIT License
└── README.md                  # This file
```

## How to Use Skills

1. **Browse the skills** in this repository
2. **Read the SKILL.md** file for each skill to understand its purpose
3. **Follow the instructions** in the skill documentation
4. **Run example scripts** to see the skill in action
5. **Adapt and customize** for your specific needs

## Adding New Skills

This repository is designed to hold multiple skills. To add a new skill:

1. Create a new directory under `skills/` with a descriptive name (use lowercase and hyphens)
2. Add a `SKILL.md` file with YAML frontmatter:
   ```yaml
   ---
   name: your-skill-name
   description: Brief description of what the skill does (max 1024 chars)
   ---
   # Instructions
   [Your detailed instructions here]
   ```
3. Optionally add:
   - `scripts/` - Code that implements the skill
   - `resources/` - Templates, data files, etc.
   - `EXAMPLES.md` - Additional examples
   - `REFERENCE.md` - Reference documentation
   - `TROUBLESHOOTING.md` - Common issues and solutions

### Skill Naming Conventions

- Use lowercase letters and hyphens: `my-skill-name`
- Keep names descriptive but concise (≤64 characters)
- Match the directory name to the skill name in SKILL.md

### Example Structure for a New Skill

```
skills/
└── my-new-skill/
    ├── SKILL.md              # Required: Metadata and instructions
    ├── EXAMPLES.md           # Optional: Usage examples
    ├── scripts/              # Optional: Implementation code
    │   └── example.py
    └── resources/            # Optional: Templates and assets
        └── template.xlsx
```

## Requirements

Different skills may have different requirements. Check each skill's documentation for specific dependencies.

Common requirements:
- Python 3.7+
- pip packages (specified in each skill)

## Contributing

Contributions are welcome! Feel free to:
- Add new skills
- Improve existing skills
- Fix bugs or typos
- Add more examples

## License

MIT License - see [LICENSE](LICENSE) file for details.

## References

- [Anthropic Skills Repository](https://github.com/anthropics/skills)
- [Claude AI Documentation](https://docs.anthropic.com/)
- [Agent Skills Specification](https://agentskills.io/)

## Author

JetSquirrel - [GitHub Profile](https://github.com/JetSquirrel)
