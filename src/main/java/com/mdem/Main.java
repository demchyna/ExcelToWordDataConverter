package com.mdem;

import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.Insets;
import javafx.scene.Group;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.text.Font;
import javafx.scene.text.Text;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.math.BigInteger;
import java.util.Iterator;

public class Main extends Application {

    private File file;
    private FileInputStream fileIn;
    private FileOutputStream fileOut;
    private final FileChooser openFileChooser = new FileChooser();
    private final FileChooser saveFileChooser = new FileChooser();

    private XSSFWorkbook workbook;
    private XSSFSheet spreadsheet;

    private XWPFDocument document;
    private XWPFTable table;
    private XWPFTableRow tableRow;
    private XWPFParagraph paragraph;
    private XWPFRun run;

    private Text label;

    private Scene createScene(final Stage primaryStage) {
        final Group root = new Group();

        Button openFile = createButton("Відкрити файл для конвертування", 250, 25, 20, 20);
        Button saveFile = createButton("Зберегти результати у файл", 250, 25, 20, 55);
        Button convert = createButton("Конвертувати", 100, 60, 280, 20);

        openFileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xls", "*.xlsx"));
        saveFileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Word File", "*.docx"));

        final ScrollPane scrollPane = new ScrollPane();
        scrollPane.setPrefWidth(360);
        scrollPane.setPrefHeight(180);
        scrollPane.setLayoutX(20);
        scrollPane.setLayoutY(100);
        scrollPane.setVbarPolicy(ScrollPane.ScrollBarPolicy.ALWAYS);
        scrollPane.setPadding(new Insets(20, 0, 0, 20));

        root.getChildren().add(openFile);
        root.getChildren().add(saveFile);
        root.getChildren().add(convert);
        root.getChildren().add(scrollPane);

        openFile.setOnAction(new EventHandler<ActionEvent>() {
            public void handle(ActionEvent actionEvent) {
                file = openFileChooser.showOpenDialog(primaryStage);
                if (file != null) {
                    try {
                        fileIn = new FileInputStream(file);
                        workbook = new XSSFWorkbook(fileIn);
                        spreadsheet = workbook.getSheetAt(0);
                    } catch (FileNotFoundException e) {
                        e.printStackTrace();
                    } catch (IOException e) {
                        e.printStackTrace();
                    } finally {
                        try {
                            fileIn.close();
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                }
            }
        });

        saveFile.setOnAction(new EventHandler<ActionEvent>() {
            public void handle(ActionEvent event) {
                file = saveFileChooser.showSaveDialog(primaryStage);
                if (file != null) {
                    try {
                        fileOut = new FileOutputStream(file);
                        document.write(fileOut);
                    } catch (IOException e) {
                        e.printStackTrace();
                    } finally {
                        try {
                            fileOut.close();
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }

                }
            }
        });

        convert.setOnAction(new EventHandler<ActionEvent>() {
            public void handle(ActionEvent event) {

                document = new XWPFDocument();
                String status = "";
                int counter = 1, index = 1;

                Iterator<Row> rowIterator = spreadsheet.iterator();
                while (rowIterator.hasNext()) {
                    Row row = (XSSFRow) rowIterator.next();
                    Iterator <Cell> cellIterator = row.cellIterator();
                    while ( cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();

                        if (cell.getCellType() == Cell.CELL_TYPE_STRING && cell.getStringCellValue().equals("name")) {

                            paragraph = document.createParagraph();
                            run = paragraph.createRun();
                            run.addCarriageReturn();
                            run.setText(row.getCell(1).getStringCellValue());
                            table = document.createTable();
                            table.setCellMargins(10, 110, 10, 110);

                            status = status + index++ + ". " + row.getCell(1).getStringCellValue() + "\n\n";
                            label = new Text(status);
                            label.setFont(new Font("Courier New", 12));

                            scrollPane.setContent(label);
                        }

                        if (cell.getCellType() == Cell.CELL_TYPE_STRING && cell.getStringCellValue().equals("title")) {
                            counter = 1;

                            tableRow = table.getRow(0);
                            tableRow.getCell(0).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(630));

                            XWPFParagraph cellParagraph;
                            cellParagraph = tableRow.getCell(0).getParagraphs().get(0);
                            setCellValue(cellParagraph, row.getCell(1).getStringCellValue());

                            for (int i=2; i<=7; i++) {
                                cellParagraph = tableRow.addNewTableCell().getParagraphs().get(0);
                                setCellValue(cellParagraph, row.getCell(i).getStringCellValue());
                            }

                            tableRow.getCell(1).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(4485));
                            tableRow.getCell(2).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(770));
                            tableRow.getCell(3).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(850));
                            tableRow.getCell(4).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(860));
                            tableRow.getCell(5).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(800));
                            tableRow.getCell(6).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1100));
                        }

                        if (cell.getCellType() == Cell.CELL_TYPE_STRING && cell.getStringCellValue().equals("item")) {
                            tableRow = table.createRow();
                            XWPFParagraph cellParagraph;

                            tableRow.getCell(0).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(630));
                            tableRow.getCell(1).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(4485));
                            tableRow.getCell(2).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(770));
                            tableRow.getCell(3).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(850));
                            tableRow.getCell(4).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(860));
                            tableRow.getCell(5).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(800));
                            tableRow.getCell(6).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1100));


                            cellParagraph = tableRow.getCell(0).getParagraphs().get(0);
                            cellParagraph.setAlignment(ParagraphAlignment.CENTER);
                            setLineSpacing(cellParagraph, 228);
                            setRun(cellParagraph.createRun(), "Arial Narrow", 10, "000000", String.valueOf(counter++), true, false);
                            tableRow.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                            cellParagraph = tableRow.getCell(1).getParagraphs().get(0);
                            cellParagraph.setAlignment(ParagraphAlignment.LEFT);
                            setLineSpacing(cellParagraph, 228);
                            setRunText(cellParagraph, "Arial Narrow", 9, "000000", row.getCell(2).getStringCellValue(), false, false);
                            tableRow.getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                            cellParagraph = tableRow.getCell(2).getParagraphs().get(0);
                            cellParagraph.setAlignment(ParagraphAlignment.CENTER);
                            setLineSpacing(cellParagraph, 228);
                            setRun(cellParagraph.createRun(), "Arial Narrow", 10, "000000", String.valueOf((int) row.getCell(3).getNumericCellValue()), false, false);
                            tableRow.getCell(2).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                            cellParagraph = tableRow.getCell(3).getParagraphs().get(0);
                            cellParagraph.setAlignment(ParagraphAlignment.CENTER);
                            setLineSpacing(cellParagraph, 228);

                            String creditsText;
                            double credits = row.getCell(4).getNumericCellValue();
                            int whole = (int) credits;
                            if((credits - whole) > 0) {
                                creditsText = String.valueOf(row.getCell(4).getNumericCellValue()).replace('.', ',');
                            } else {
                                creditsText = String.valueOf(whole);
                            }

                            setRun(cellParagraph.createRun(), "Arial Narrow", 10, "000000", creditsText, false, false);
                            tableRow.getCell(3).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                            cellParagraph = tableRow.getCell(4).getParagraphs().get(0);
                            cellParagraph.setAlignment(ParagraphAlignment.CENTER);
                            setLineSpacing(cellParagraph, 228);
                            setRun(cellParagraph.createRun(), "Arial Narrow", 10, "000000", String.valueOf((int) row.getCell(5).getNumericCellValue()), false, false);
                            tableRow.getCell(4).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                            cellParagraph = tableRow.getCell(5).getParagraphs().get(0);
                            cellParagraph.setAlignment(ParagraphAlignment.CENTER);
                            setLineSpacing(cellParagraph, 228);
                            setRun(cellParagraph.createRun(), "Arial Narrow", 10, "000000", row.getCell(6).getStringCellValue(), false, false);
                            tableRow.getCell(5).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                            cellParagraph = tableRow.getCell(6).getParagraphs().get(0);
                            cellParagraph.setAlignment(ParagraphAlignment.CENTER);
                            setLineSpacing(cellParagraph, 228);
                            setRunText(cellParagraph, "Arial Narrow", 9, "000000", row.getCell(7).getStringCellValue(), false, false);
                            tableRow.getCell(6).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                        }

                        if (cell.getCellType() == Cell.CELL_TYPE_STRING && cell.getStringCellValue().equals("section")) {
                            counter = 1;

                            XWPFParagraph cellParagraph;
                            tableRow = table.createRow();
                            int numberOfRows = table.getNumberOfRows();

                            tableRow.getCell(0).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(630));
                            tableRow.getCell(1).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(4485));
                            tableRow.getCell(2).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(770));
                            tableRow.getCell(3).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(850));
                            tableRow.getCell(4).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(860));
                            tableRow.getCell(5).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(800));
                            tableRow.getCell(6).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1100));

                            cellParagraph = tableRow.getCell(0).getParagraphs().get(0);
                            cellParagraph.setAlignment(ParagraphAlignment.LEFT);
                            setLineSpacing(cellParagraph, 228);
                            setRun(cellParagraph.createRun(), "Arial Narrow", 10, "000000", "", true, false);
                            tableRow.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                            cellParagraph = tableRow.getCell(1).getParagraphs().get(0);
                            cellParagraph.setAlignment(ParagraphAlignment.LEFT);
                            cellParagraph.setIndentationLeft(-80);
                            cellParagraph.setIndentationRight(-80);
                            setLineSpacing(cellParagraph, 228);
                            setRunText(cellParagraph, "Arial Narrow", 10, "000000", row.getCell(1).getStringCellValue(), true, false);
                            tableRow.getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                            mergeCellHorizontally(table, numberOfRows-1, 1, 6);
                        }
                    }
                }
            }
        });

        return new Scene(root, 400, 300);
    }

    private  Button createButton(String text, double width, double height, double posX, double posY) {
        Button button = ButtonBuilder.create()
                .text(text)
                .prefWidth(width)
                .prefHeight(height)
                .layoutX(posX)
                .layoutY(posY)
                .build();
        return button;
    }

    private Label createLabel(String text, double width, double height, double posX, double posY) {
        Label label = LabelBuilder.create()
                .text(text)
                .prefWidth(width)
                .prefHeight(height)
                .layoutX(posX)
                .layoutY(posY)
                .build();
        return label;
    }

    @Override
    public void start(Stage primaryStage) throws Exception {
        primaryStage.setTitle("Rating from Excel to Word");
        primaryStage.setScene(createScene(primaryStage));
        primaryStage.show();
    }

    public void setCellValue(XWPFParagraph cellParagraph, String value) {
        cellParagraph.setAlignment(ParagraphAlignment.CENTER);
        cellParagraph.setIndentationLeft(-85);
        cellParagraph.setIndentationRight(-85);
        setLineSpacing(cellParagraph, 228);
        setRunText(cellParagraph, "Arial Narrow", 10, "000000", value, true, false);
    }

    public static void setLineSpacing(XWPFParagraph cellParagraph, long space) {
        CTPPr ppr = cellParagraph.getCTP().getPPr();
        if (ppr == null) ppr = cellParagraph.getCTP().addNewPPr();
        CTSpacing spacing = ppr.isSetSpacing()? ppr.getSpacing() : ppr.addNewSpacing();
        spacing.setAfter(BigInteger.valueOf(0));
        spacing.setBefore(BigInteger.valueOf(0));
        spacing.setLineRule(STLineSpacingRule.AUTO);
        spacing.setLine(BigInteger.valueOf(space));
    }

    public static void setRunText (XWPFParagraph cellParagraph , String fontFamily , int fontSize , String colorRGB , String text , boolean bold , boolean addBreak) {

        int slashIndex = 0;
        if (text.contains("Rating points")) {
            slashIndex = 15;
        } else {
            slashIndex = text.lastIndexOf(47);
        }

        if (text.contains("національною")) {
            String part1, part2;

            part1 = text.substring(0, text.indexOf('а', 12)+1);
            part2 = text.substring(text.indexOf('а', 12)+1);
            text = part1 + "\n" + part2;
        }

        String ukrText = text.substring(0, slashIndex+1);
        String engText = text.substring(slashIndex+1);

        XWPFRun run = cellParagraph.createRun();
        run.setFontFamily(fontFamily);
        run.setFontSize(fontSize);
        run.setColor(colorRGB);
        run.setBold(bold);
        run.setItalic(false);
        run.setText(ukrText);

        run = cellParagraph.createRun();
        run.setFontFamily(fontFamily);
        run.setFontSize(fontSize);
        run.setColor(colorRGB);
        run.setBold(bold);
        run.setItalic(true);
        if (engText.equals(" Excellent") || engText.equals(" Good") || engText.equals(" Satisfactory")) {
            run.setText(" ");
            run.addBreak();
            run.setText(engText.substring(1));
        } else {
            run.setText(engText);
        }
    }


    public static void setRun (XWPFRun run , String fontFamily , int fontSize , String colorRGB , String text , boolean bold , boolean addBreak) {
        run.setFontFamily(fontFamily);
        run.setFontSize(fontSize);
        run.setColor(colorRGB);
        run.setText(text);
        run.setBold(bold);
        run.setItalic(false);
    }

    public static void mergeCellHorizontally(XWPFTable table, int row, int fromCol, int toCol) {
        for(int colIndex = fromCol; colIndex <= toCol; colIndex++){
            CTHMerge hmerge = CTHMerge.Factory.newInstance();
            if(colIndex == fromCol){
                // The first merged cell is set with RESTART merge value
                hmerge.setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                hmerge.setVal(STMerge.CONTINUE);
            }
            XWPFTableCell cell = table.getRow(row).getCell(colIndex);
            // Try getting the TcPr. Not simply setting an new one every time.
            CTTcPr tcPr = cell.getCTTc().getTcPr();
            if (tcPr != null) {
                tcPr.setHMerge(hmerge);
            } else {
                // only set an new TcPr if there is not one already
                tcPr = CTTcPr.Factory.newInstance();
                tcPr.setHMerge(hmerge);
                cell.getCTTc().setTcPr(tcPr);
            }
        }
    }

    public static void main(String[] args) {
        launch(args);
    }
}
