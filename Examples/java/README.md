# Java Example - DWG to PDF Conversion

## Prerequisites

1. Register the COM DLL (run as Administrator):
   - 64-bit: `regsvr32 "path\to\dwgtopdfx64.dll"`
   - 32-bit: `regsvr32 "path\to\dwgtopdfx.dll"`

2. Java JDK 8+
3. [Jacob (Java COM Bridge)](https://github.com/freemansoft/jacob-project) library:
   - Download `jacob.jar` and place it in the `lib/` folder.
   - Download `jacob-<arch>.dll` and place it in `C:\Windows\System32\` (64-bit) or `C:\Windows\SysWOW64\` (32-bit).

4. Place your test DWG file in this directory or adjust the file paths.

## How to Build and Run

```
mkdir lib
# Copy jacob.jar to lib/

javac -cp "lib\jacob.jar" src\Example.java -d out
java -cp "out;lib\jacob.jar" -Djava.library.path="C:\Windows\System32" Example
```

## Examples Included

The file contains 4 methods:

1. `basicConversion()` - Simple DWG to PDF
2. `onePdfPerView()` - Export each view/layout
3. `batchConversion()` - Convert multiple files
4. `advancedFeatures()` - Full API usage

## Notes

- Uses Jacob's `ActiveXComponent` to create the COM object.
- The Jacob DLL architecture must match the JVM architecture.