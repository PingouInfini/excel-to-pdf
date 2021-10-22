# ExcelToPDF

### Prerequisite for IntelliJ

1. Open project structure (ctrl-shift-alt-S)
2. Tab Librairies, add the "lib" path of javafx (ie:"C:\Program Files\Java\javafx-sdk-17.0.0.1\lib")
3. Edit configuration of "Main" and add this for VM options:
```
--module-path "C:\Program Files\Java\javafx-sdk-17.0.0.1\lib" --add-modules javafx.controls,javafx.fxml
```

### Build Project

1. Open project structure (ctrl-shift-alt-S)
2. Tab Artifact, add Jar > From module with dep
3. Select main Class, then ok, ok
4. Build > build Artifacts > Build

Check if build is OK by running App in standalone
Go in artifact path "F:\workspace\ExcelToPDF\out\artifacts\ExcelToPDF_jar", then run
```
java --module-path "C:\Program Files\Java\javafx-sdk-17.0.0.1\lib" --add-modules javafx.controls,javafx.fxml -jar ExcelToPDF.jar 
```


set PATH_TO_FX="C:\Program Files\Java\javafx-sdk-17.0.0.1\lib"
javac --module-path %PATH_TO_FX% --add-modules javafx.controls Main.java
javac --module-path "C:\Program Files\Java\javafx-sdk-17.0.0.1\lib" --add-modules javafx.controls,javafx.fxml Main.java
java --module-path %PATH_TO_FX% --add-modules javafx.controls Main

### How to create exe from source

```
https://medium.com/@vinayprabhu19/creating-executable-javafx-application-part-2-c98cfa65801e
```