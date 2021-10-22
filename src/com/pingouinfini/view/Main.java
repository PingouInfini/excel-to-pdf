package com.pingouinfini.view;


import com.pingouinfini.controller.CreateExcelFileTask;
import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Separator;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.VBox;
import javafx.scene.text.Text;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.eadge.extractpdfexcel.PdfConverter;
import org.eadge.extractpdfexcel.data.ExtractedData;
import org.eadge.extractpdfexcel.data.SortedData;
import org.eadge.extractpdfexcel.data.XclPage;
import org.eadge.extractpdfexcel.data.block.Block;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

public class Main extends Application {

    String PATTERN_DATE = "^(?:(?:31(\\/|-|\\.)(?:0?[13578]|1[02]))\\1|(?:(?:29|30)(\\/|-|\\.)(?:0?[13-9]|1[0-2])\\2))(?:(?:1[6-9]|[2-9]\\d)?\\d{2})$|^(?:29(\\/|-|\\.)0?2\\3(?:(?:(?:1[6-9]|[2-9]\\d)?(?:0[48]|[2468][048]|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00))))$|^(?:0?[1-9]|1\\d|2[0-8])(\\/|-|\\.)(?:(?:0?[1-9])|(?:1[0-2]))\\4(?:(?:1[6-9]|[2-9]\\d)?\\d{2})$";
    Text status = new Text();

    private CreateExcelFileTask createExcelFileTask;

    @Override
    public void start(Stage primaryStage) {

        // Filechooser pour N fichiers
        final FileChooser fileChooser = new FileChooser();
        configuringFileChooser(fileChooser);

        // DirectoryChooser pour 1 répertoire
        DirectoryChooser directoryChooser = new DirectoryChooser();
        configuringDirectoryChooser(directoryChooser);

        // Button pour le Filechooser
        Image imagePdf = new Image(getClass().getResourceAsStream("/icon/icon-pdf.png"));
        Button buttonpdf = new Button("Transformer un ou plusieurs fichier(s) PDF en fichier(s) Excel", new ImageView(imagePdf));

        // Button pour le DirectoryChooser
        Image imageFolder = new Image(getClass().getResourceAsStream("/icon/icon-folder.png"));
        Button buttonFolder = new Button("Transformer les fichiers PDF d'un répertoire en fichiers Excel", new ImageView(imageFolder));

        // Statut
        final Separator separatorup = new Separator();
        final Separator separatordown = new Separator();

        Text blankspace = new Text();

        status.setText("Veuillez sélectionner une option");

        //ProgressIndicator progressIndicator = new ProgressIndicator();
        //ProgressBar progressBar = new ProgressBar();

        buttonpdf.setOnAction(event -> {
            List<File> files = fileChooser.showOpenMultipleDialog(primaryStage);

            status.setText("Transformation d'un ou plusieurs fichier(s) PDF en fichier(s) Excel");
            for (File file : files) {
                System.out.println(file.getName());
                status.setText("Traitement fichier : " + file.getName());
                //buttonpdf.setDisable(true);
                //progressBar.setProgress(0);
                //progressIndicator.setProgress(0);

                // Create a Task.
                //createExcelFileTask = new CreateExcelFileTask();
                createExcelFile(file);
            }
        });

        buttonFolder.setOnAction(event -> {
            File selectedDirectory = directoryChooser.showDialog(primaryStage);

            transformFilesInDirectory(selectedDirectory);
        });

        VBox root = new VBox();
        root.setPadding(new Insets(10));
        root.setSpacing(5);

        root.getChildren().addAll(buttonpdf, buttonFolder, blankspace, separatorup, status, separatordown);

        Scene scene = new Scene(root, 400, 200);

        primaryStage.setTitle("From PDF to Excel");
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    private void transformFilesInDirectory(File file) {
        for (File currentFile : Objects.requireNonNull(file.listFiles())) {

            if (currentFile.isDirectory())
                transformFilesInDirectory(currentFile);

            String extension = getExtension(currentFile.getName()).isPresent() ? getExtension(currentFile.getName()).get() : "";
            if (!"pdf".equals(extension))
                continue;
            createExcelFile(currentFile);
        }
    }

    private void createExcelFile(File file) {
        String filename = file.getName();
        String filePath = file.getAbsolutePath();
        String fileDirectory = file.getParent();
        String fileNameWithoutExtension = filename.substring(0, filename.length() - 4);
        String excelFilePAth = fileDirectory + "\\" + fileNameWithoutExtension + ".xls";

        try {
            ExtractedData extractedData = PdfConverter.extractFromFile(filePath);
            SortedData sortedData = PdfConverter.sortExtractedData(extractedData, 0, 1);
            ArrayList<XclPage> excelPages = PdfConverter.createExcelPages(sortedData);

            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheetExtract = workbook.createSheet("Données brutes");
            HSSFSheet sheetTab = workbook.createSheet("Données filtrées");

            for (XclPage xclPage : excelPages) {
                customCreateExcelSheet(sheetExtract, xclPage, 0.0D, 0.0D);
                customCreateExcelSheet(sheetTab, xclPage, 0.0D, 0.0D, PATTERN_DATE);
            }
            PdfConverter.createExcelPages(sortedData);
            FileOutputStream out = new FileOutputStream(excelFilePAth);
            workbook.write(out);
            out.close();
        } catch (IOException e1) {
            e1.printStackTrace();
        }

    }

    private void customCreateExcelSheet(HSSFSheet sheet, XclPage xclPage, double lineFactor, double columnFactor) {
        int lastLineNb = sheet.getLastRowNum() == 0 ? sheet.getLastRowNum() : sheet.getLastRowNum() + 1;

        for (int line = 0; line < xclPage.numberOfLines(); ++line) {
            HSSFRow createdLine = sheet.createRow(lastLineNb + line);
            if (lineFactor != 0.0D) {
                createdLine.setHeight((short) ((int) (xclPage.getLineHeight(line) * 50.0D)));
            }

            for (int col = 0; col < xclPage.numberOfColumns(); ++col) {
                Block block = xclPage.getBlockAt(col, line);
                if (block != null) {
                    setCellValue(createdLine, col, block);
                }
            }
        }

        if (columnFactor != 0.0D) {
            for (int col = 0; col < xclPage.numberOfColumns(); ++col) {
                sheet.setColumnWidth(col, (int) xclPage.getColumnWidth(col) * 20);
            }
        }
    }

    private void setCellValue(HSSFRow createdLine, int col, Block block) {
        HSSFCell createdCell = createdLine.createCell(col);

        try {
            // force check if its a number
            Float.parseFloat(block.getFormattedText());
            createdCell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
            createdCell.setCellValue(Float.parseFloat(block.getFormattedText()));
        } catch (NumberFormatException e) {
            //not float
            createdCell.setCellValue(block.getFormattedText());
        }
    }

    private void customCreateExcelSheet(HSSFSheet sheet, XclPage xclPage, double lineFactor, double columnFactor, String filtredPattern) {
        boolean init = false;
        int lastLineNb = sheet.getLastRowNum() + 1;

        // La sheet est vide, on passe en mode 'init' pour rajouter la ligne d'entete
        // qui est la ligne précédent la premiere fois ou on rencontrera le pattern
        if (sheet.getLastRowNum() == 0) {
            init = true;
            lastLineNb = sheet.getLastRowNum();
        }

        for (int line = 0; line < xclPage.numberOfLines(); ++line) {

            // Check if 1st cell of this line is not null and match with the pattern
            // else go to the next line
            Block checkBlock = xclPage.getBlockAt(0, line);
            if (checkBlock == null || !checkBlock.getFormattedText().matches(filtredPattern)) {
                lastLineNb--;
                continue;
            }

            HSSFRow createdLine = sheet.createRow(lastLineNb + line);
            if (lineFactor != 0.0D) {
                createdLine.setHeight((short) ((int) (xclPage.getLineHeight(line) * 50.0D)));
            }

            for (int col = 0; col < xclPage.numberOfColumns(); ++col) {
                Block block = xclPage.getBlockAt(col, line);

                // Si on arrive ici, c'est que la 1er cellule de la ligne match avec le pattern
                if (block != null) {
                    // Si on est en phase d'init, on va recup la ligne d'entete, puis on rajoute la ligne
                    if (init) {
                        // get all col for previous line
                        for (int colprev = 0; colprev < xclPage.numberOfColumns(); ++colprev) {
                            Block blockinit = xclPage.getBlockAt(colprev, line - 1);
                            if (blockinit != null) {
                                createdLine.createCell(colprev).setCellValue(blockinit.getFormattedText());
                            }
                        }
                        lastLineNb++;
                        createdLine = sheet.createRow(lastLineNb + line);
                        init = false;
                    }
                    setCellValue(createdLine, col, block);
                }
            }
        }

        if (columnFactor != 0.0D) {
            for (int col = 0; col < xclPage.numberOfColumns(); ++col) {
                sheet.setColumnWidth(col, (int) xclPage.getColumnWidth(col) * 20);
            }
        }
    }

    private void configuringDirectoryChooser(DirectoryChooser directoryChooser) {
        // Set title for DirectoryChooser
        directoryChooser.setTitle("Sélectionner un répertoire");

        // Set Initial Directory
        directoryChooser.setInitialDirectory(new File(System.getProperty("user.home")));
    }


    private void configuringFileChooser(FileChooser fileChooser) {
        // Set title for FileChooser
        fileChooser.setTitle("Sélectionner un ou plusieurs fichiers");

        // Set Initial Directory
        fileChooser.setInitialDirectory(new File(System.getProperty("user.home")));

        // Add Extension Filters
        fileChooser.getExtensionFilters().addAll(//
                new FileChooser.ExtensionFilter("PDF", "*.pdf"));
    }

    private Optional<String> getExtension(String filename) {
        return Optional.ofNullable(filename)
                .filter(f -> f.contains("."))
                .map(f -> f.substring(filename.lastIndexOf(".") + 1));
    }

    public static void main(String[] args) {
        launch(args);
    }

}
