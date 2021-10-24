# ExcelToPDF


### Build Project

1. Open project structure (ctrl-shift-alt-S)
2. Tab Artifact, add Jar > From module with dep
3. Select main Class, then ok, ok
4. Build > build Artifacts > Build

Check if build is OK by running App in standalone
Go in artifact path "F:\workspace\ExcelToPDF\out\artifacts\ExcelToPDF_jar", then run
```
java -jar ExcelToPDF.jar 
```

### How to create exe from source

Using launch4j

0. (if exists) Unzip ExcelToPDF\target\jre\lib\modules.zip
0. (if exists) Remove 'modules.zip'
1. Load configuration using file 'l4jconfig.xml'
2. Make sure 'target/jre' directory is present
3. Click on the wheel