package com.pingouinfini;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.eadge.extractpdfexcel.PdfConverter;
import org.eadge.extractpdfexcel.data.ExtractedData;
import org.eadge.extractpdfexcel.data.SortedData;
import org.eadge.extractpdfexcel.data.XclPage;
import org.eadge.extractpdfexcel.data.block.Block;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.border.LineBorder;
import javax.swing.filechooser.FileSystemView;
import java.awt.*;
import java.beans.PropertyChangeEvent;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;


public class ExcelToPdf {
    private String currentFileName = "";

    public static void main(String[] args) {
        new ExcelToPdf();
    }

    private ExcelToPdf() {
        EventQueue.invokeLater(() -> {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | UnsupportedLookAndFeelException ignored) {
            }

            JFrame frame = new JFrame("PDF to Excel");
            frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            frame.setLayout(new BorderLayout());
            frame.add(new MainPane());
            frame.pack();
            frame.setLocationRelativeTo(null);
            frame.setVisible(true);
        });
    }

    public class MainPane extends JPanel {
        private JProgressBar pbProgress;
        private JButton buttonFile, buttonFolder;
        private JTextField status;


        MainPane() {

            setBorder(new EmptyBorder(10, 10, 10, 10));
            setLayout(new GridBagLayout());
            GridBagConstraints gbc = new GridBagConstraints();
            gbc.insets = new Insets(4, 4, 4, 4); // marges
            gbc.gridx = gbc.gridy = 0; // la grille commence en (0, 0)
            gbc.gridwidth = GridBagConstraints.REMAINDER; // seul composant de sa colonne, il est donc le dernier.
            gbc.gridheight = 1; // valeur par défaut - peut s'étendre sur une seule ligne.

            gbc.anchor = GridBagConstraints.BASELINE_LEADING; // LINE_START ou BASELINE_LEADING mais pas WEST.
            gbc.fill = GridBagConstraints.BOTH;

            // Button pour le Filechooser
            Icon iconPdf = new ImageIcon(getClass().getResource("/icon/icon-pdf.png"));
            buttonFile = new JButton("Transformer un ou plusieurs fichier(s) PDF en fichier(s) Excel", iconPdf);
            gbc.gridy++;
            add(buttonFile, gbc);

            buttonFile.addActionListener(e -> {
                buttonFile.setEnabled(false);
                buttonFolder.setEnabled(false);
                PDFFilesWorker fileWorker = new PDFFilesWorker();
                fileWorker.addPropertyChangeListener(evt -> evtProgressBar(evt, buttonFile, buttonFolder));
                fileWorker.execute();
            });

            // Button pour le DirectoryChooser
            Icon iconFolder = new ImageIcon(getClass().getResource("/icon/icon-folder.png"));
            buttonFolder = new JButton("Transformer les fichiers PDF d'un répertoire en fichiers Excel", iconFolder);
            gbc.gridy++;
            add(buttonFolder, gbc);

            buttonFolder.addActionListener(e -> {
                buttonFolder.setEnabled(false);
                buttonFile.setEnabled(false);
                PDFFolderWorker folderWorker = new PDFFolderWorker();
                folderWorker.addPropertyChangeListener(evt -> evtProgressBar(evt, buttonFile, buttonFolder));
                folderWorker.execute();
            });

            JSeparator separator = new JSeparator();
            gbc.gridy++;
            add(separator, gbc);

            JLabel label = new JLabel("Statut :");
            gbc.gridy++;
            add(label, gbc);

            pbProgress = new JProgressBar();
            pbProgress.setStringPainted(true);
            gbc.gridy++;
            add(pbProgress, gbc);

            status = new JTextField();
            status.setText("init");
            status.setEnabled(false);
            status.setBorder(new LineBorder(Color.GRAY, 1));
            status.setBackground(Color.LIGHT_GRAY);
            gbc.gridy++;
            add(status, gbc);

        }

        private void evtProgressBar(PropertyChangeEvent evt, JButton buttonFile, JButton buttonForder) {
            String name = evt.getPropertyName();
            if (name.equals("progress")) {
                int progress = (int) evt.getNewValue();
                pbProgress.setValue(progress);
                status.setText(currentFileName);
                repaint();
            } else if (name.equals("state")) {
                SwingWorker.StateValue state = (SwingWorker.StateValue) evt.getNewValue();
                if (state == SwingWorker.StateValue.DONE) {
                    buttonFile.setEnabled(true);
                    buttonForder.setEnabled(true);

                    status.setText("Terminé !");
                }
            }
        }
    }


    public class PDFFilesWorker extends SwingWorker<Object, Object> {
        @Override
        protected Object doInBackground() {
            FileDialog fd = new FileDialog((Frame) null, "Sélectionnez un ou plusieurs fichier(s)", FileDialog.LOAD);
            fd.setMultipleMode(true);
            fd.setDirectory(FileSystemView.getFileSystemView().getHomeDirectory().getAbsolutePath());
            fd.setFile("*.pdf");
            fd.setVisible(true);

            File[] files = fd.getFiles();
            int i = 1;
            for (File file : files) {
                createExcelFile(file);
                currentFileName = file.getName();
                setProgress(i * 100 / (files.length));
                i++;
            }
            return null;
        }
    }

    public class PDFFolderWorker extends SwingWorker<Object, Object> {
        @Override
        protected Object doInBackground() {
            JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
            jfc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

            int returnValue = jfc.showOpenDialog(null);
            if (returnValue == JFileChooser.APPROVE_OPTION) {
                File selectedDirectory = jfc.getSelectedFile();

                Set<File> allPdfFiles = getAllFilesInDirectory(selectedDirectory);
                int i = 1;
                for (File file : allPdfFiles) {
                    createExcelFile(file);
                    currentFileName = file.getName();
                    setProgress(i * 100 / (allPdfFiles.size()));
                    i++;
                }
            }
            return null;
        }
    }

    private Set<File> getAllFilesInDirectory(File file) {
        Set<File> allPdfFiles = new HashSet<>();
        for (File currentFile : Objects.requireNonNull(file.listFiles())) {

            if (currentFile.isDirectory())
                allPdfFiles.addAll(getAllFilesInDirectory(currentFile));

            String extension = getExtension(currentFile.getName()).isPresent() ? getExtension(currentFile.getName()).get() : "";
            if (!"pdf".equalsIgnoreCase(extension))
                continue;
            allPdfFiles.add(currentFile);
        }
        return allPdfFiles;
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
                String PATTERN_DATE = "^(?:(?:31(\\/|-|\\.)(?:0?[13578]|1[02]))\\1|(?:(?:29|30)(\\/|-|\\.)(?:0?[13-9]|1[0-2])\\2))(?:(?:1[6-9]|[2-9]\\d)?\\d{2})$|^(?:29(\\/|-|\\.)0?2\\3(?:(?:(?:1[6-9]|[2-9]\\d)?(?:0[48]|[2468][048]|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00))))$|^(?:0?[1-9]|1\\d|2[0-8])(\\/|-|\\.)(?:(?:0?[1-9])|(?:1[0-2]))\\4(?:(?:1[6-9]|[2-9]\\d)?\\d{2})$";
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

    private Optional<String> getExtension(String filename) {
        return Optional.ofNullable(filename)
                .filter(f -> f.contains("."))
                .map(f -> f.substring(filename.lastIndexOf(".") + 1));
    }
}
