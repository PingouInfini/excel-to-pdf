# excel-to-pdf


### Build Project

1. Open project structure (ctrl-shift-alt-S)
2. Tab Artifact, add Jar > From module with dep
3. Select main Class, then ok, ok
4. Build > build Artifacts > Build

Check if build is OK by running App in standalone
Go in artifact path "F:\workspace\excel-to-pdf\out\artifacts\ExcelToPDF_jar", then run
```
java -jar ExcelToPDF.jar 
```

### How to create exe from source

Using launch4j

0. (if exists) Unzip excel-to-pdf\target\jre\lib\modules.zip with 'extract here' (DO NOT EXTRACT AS FOLDER !)
0. (if exists) Remove 'modules.zip'
1. Load configuration using file 'l4jconfig.xml'
2. Make sure 'target/jre' directory is present
3. Click on the wheel