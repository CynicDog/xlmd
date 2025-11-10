# Example: Excel to Markdown
```
./gradlew runApp --args="-i src/test/resources/sample/iris/in.xlsx -o src/test/resources/sample/iris/out.md"
```

# Example: Markdown to Excel
```
./gradlew runApp --args="-i src/test/resources/sample/sales/in.md -o src/test/resources/sample/sales/out.xlsx"
```

#### 2. Building Executables (CI/CD)

The packaging process uses `jpackage` to create a bundled executable that includes all your code and a private Java runtime, so users don't need to install Java.

1.  **Build the JAR (If needed, but packaging tasks do this automatically):**
    ```bash
    ./gradlew jar 
    # Output: build/libs/xlmd-app.jar

2.  **Build Windows Executable (`.exe`):** (Run on a Windows runner)
    ```bash
    ./gradlew packageWindows 
    # Output: Executable files in build/executables/

3.  **Build Unix Executable:** (Run on a Linux or macOS runner)
    ```bash
    ./gradlew packageUnix 
    # Output: Application image directory (e.g., build/executables/xlmd/)