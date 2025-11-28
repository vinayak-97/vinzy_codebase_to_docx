# Codebase to DOCX

Convert your entire codebase to a Word document with ease! This Python package scans your project directory and creates a well-formatted Word document containing all your code files.

## Features

- 📁 **Automatic file discovery** - Recursively scans your codebase
- 📝 **File path inclusion** - Optional display of full file paths
- 📑 **Table of contents** - Easy navigation through your codebase
- 🚫 **Smart filtering** - Automatically ignores common directories (node_modules, .git, etc.)
- 🎨 **Clean formatting** - Code presented in monospace font with clear headings
- ⚙️ **Customizable** - Configure which files and directories to include

## Installation

```bash
pip install vinzy_codebase_to_docx
```

## Usage

### Basic Usage

```python
from vinzy_codebase_to_docx import convert_codebase

# Convert entire codebase
convert_codebase(
    codebase_path="/path/to/your/project",
    output_path="my_codebase.docx"
)
```

### Advanced Usage

```python
from vinzy_codebase_to_docx import CodebaseConverter

converter = CodebaseConverter(
    codebase_path="/path/to/your/project",
    output_path="my_codebase.docx",
    include_file_paths=True,  # Show file locations
    include_toc=True,          # Include table of contents
    ignore_dirs={'tests', 'docs'},  # Additional dirs to ignore
    include_extensions={'.py', '.js'}  # Only include these files
)

# Convert with progress tracking
def progress(current, total, filename):
    print(f"Processing {current}/{total}: {filename}")

converter.convert(progress_callback=progress)
```

### Command Line Usage

```bash
python codebase_to_docx.py /path/to/codebase output.docx
```

## Configuration Options

| Parameter            | Type     | Default           | Description                      |
| -------------------- | -------- | ----------------- | -------------------------------- |
| `codebase_path`      | str      | Required          | Path to your codebase directory  |
| `output_path`        | str      | `'codebase.docx'` | Output Word document path        |
| `include_file_paths` | bool     | `True`            | Show full file paths in document |
| `include_toc`        | bool     | `True`            | Include table of contents        |
| `ignore_dirs`        | Set[str] | See below         | Directories to skip              |
| `include_extensions` | Set[str] | See below         | File extensions to include       |

### Default Ignored Directories

```python
{
    '__pycache__', 'node_modules', '.git', '.venv', 'venv',
    'env', 'build', 'dist', '.idea', '.vscode', 'target',
    '.gradle', 'bin', 'obj', 'coverage'
}
```

### Default Included Extensions

```python
{
    '.py', '.js', '.jsx', '.ts', '.tsx', '.java', '.cpp', '.c', '.h',
    '.cs', '.rb', '.go', '.rs', '.php', '.swift', '.kt', '.scala',
    '.html', '.css', '.scss', '.sass', '.less', '.xml', '.json',
    '.yaml', '.yml', '.md', '.txt', '.sh', '.bash', '.sql', '.r'
}
```

## Examples

### Example 1: Convert Python Project Only

```python
from codebase_to_docx import convert_codebase

convert_codebase(
    codebase_path="/path/to/project",
    output_path="python_code.docx",
    include_extensions={'.py'}
)
```

### Example 2: Include Tests but Ignore Documentation

```python
from codebase_to_docx import CodebaseConverter

converter = CodebaseConverter(
    codebase_path="/path/to/project",
    output_path="codebase_with_tests.docx",
    ignore_dirs={'docs', 'documentation'}
)
converter.convert()
```

### Example 3: Custom Extensions for Web Project

```python
convert_codebase(
    codebase_path="/path/to/web/project",
    output_path="web_project.docx",
    include_extensions={'.html', '.css', '.js', '.jsx', '.tsx'}
)
```

## Use Cases

- 📚 **Documentation** - Create comprehensive code documentation
- 🔍 **Code Review** - Share entire codebase for review
- 📦 **Archival** - Archive project snapshots
- 🎓 **Education** - Share code examples with students
- 💼 **Presentations** - Include code in professional documents

## Requirements

- Python 3.7+
- python-docx>=0.8.11

## License

MIT License

## Tips

1. **Large Codebases**: For very large projects, consider filtering by specific directories or file types
2. **Binary Files**: The converter automatically skips binary files and handles encoding errors
3. **Performance**: Processing time depends on codebase size; expect ~100-500 files per minute
4. **File Size**: Large codebases may result in large Word documents; consider splitting if needed

## Troubleshooting

**Issue**: Out of memory error

- **Solution**: Process smaller portions of your codebase or increase available memory

**Issue**: Some files not appearing

- **Solution**: Check if they're in ignored directories or have unsupported extensions

**Issue**: Encoding errors

- **Solution**: The converter uses `errors='ignore'` to handle encoding issues automatically
