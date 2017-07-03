# VSFormatter

VSFormatter is a rudimentary auto-formatter for Visual Studio.

It creates an instance of Visual Studio, opens a set of *.cs files and
auto-formats them using SelectAll() + SmartFormat(), saves the file(s)
and then quits VS without saving the solution.

Current version references and targets VS14 (Visual Studio 15), but can
easily be changed to target any version of VS by removing and re-adding
the "Microsoft Development Environment 8.0" COM reference and replacing
the version number in the "VisualStudio.DTE.14.0" ProgID reference

### Sample usage:

```
C:\> VSFormatter.exe C:\path\to\cs\files
```

VSFormatter will recursively search for files with the `.cs` extension in the directory `C:\path\to\cs\files` and autoformat any syntactically valid file.
