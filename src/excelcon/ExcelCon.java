/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelcon;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TextArea;
import javafx.scene.effect.BlendMode;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.StackPane;
import javafx.scene.paint.Color;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.FileChooser.ExtensionFilter;
import javafx.stage.Stage;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author ali
 */
public class ExcelCon extends Application {

    static FileInputStream xls = null;
    File selectedDirectory;

    @Override
    public void start(Stage primaryStage) throws FileNotFoundException, IOException {
        Button btn = new Button();
        Button saveB = new Button();
        saveB.setText("Save");
        Button run = new Button();
        run.setText("Run");
        TextArea t = new TextArea();
        TextArea saveP = new TextArea();
        t.setMaxHeight(1);
        t.setWrapText(true);
        t.setMaxWidth(300);
        saveP.setMaxHeight(1);
        saveP.setWrapText(true);
        saveP.setMaxWidth(300);
        Label into = new Label();
        Label don = new Label();
        btn.setText("Browse");
        FileChooser chooser = new FileChooser();
        chooser.setTitle("Choose");
        FileChooser.ExtensionFilter ef = new ExtensionFilter("xlsx", "*.xlsx");
        chooser.getExtensionFilters().add(ef);
        GridPane gridpane = new GridPane();
        gridpane.setPadding(new Insets(5));
        gridpane.setHgap(10);
        gridpane.setVgap(10);
        into.setText("Start by choosing \"xlsx\" file , press Browse\n "
                + "{TMX,ABO_Open,TMX_ABO_Open"
                + ",\nVac.,Caps,Inhibit,Sensor,Sensor_Clip,Color}");
        t.setText("File Name");
        DirectoryChooser dirCh = new DirectoryChooser();
        dirCh.setTitle("Save Files");

        //System.out.println(selectedDirectory.getPath());
        String[][] names = new String[10][2];
        names[0][0] = "TMX";
        names[0][1] = "TMX_Of:";
        names[1][0] = "ABO_Open";
        names[2][0] = "ABO";
        names[3][0] = "Vac.";
        names[4][0] = "Sensor_Clip";
        names[5][0] = "Color";
        names[6][0] = "Sensor";
        names[7][0] = "TMX_ABO_Open";
        names[8][0] = "Caps";
        names[9][0] = "Inhibit";
        names[1][1] = "ABO_Open_Of:";
        names[2][1] = "ABO_Of:";
        names[3][1] = "Vac._Of:";
        names[4][1] = "Sensor_Clip_Of:";
        names[5][1] = "Color_Of:";
        names[6][1] = "Sensor_Of:";
        names[7][1] = "TMX_ABO_Open_Of:";
        names[8][1] = "Caps_Of:";
        names[9][1] = "Inhibit_Of:";
        btn.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                try {
                    File file = chooser.showOpenDialog(new Stage());

                    xls = new FileInputStream(file);

                    t.setText(file.getPath());

                } catch (FileNotFoundException ex) {
                    Logger.getLogger(ExcelCon.class.getName()).log(Level.SEVERE, null, ex);
                } catch (IOException ex) {
                    Logger.getLogger(ExcelCon.class.getName()).log(Level.SEVERE, null, ex);
                } finally {
                    try {
                        xls.close();
                    } catch (IOException ex) {
                        Logger.getLogger(ExcelCon.class.getName()).log(Level.SEVERE, null, ex);
                    }
                }
            }
        });
        saveB.setOnAction(new EventHandler<ActionEvent>() {

            @Override
            public void handle(ActionEvent event) {
                selectedDirectory = dirCh.showDialog(primaryStage);
                saveP.setText(selectedDirectory.getPath());
            }

        });

        run.setOnAction(new EventHandler<ActionEvent>() {

            @Override
            public void handle(ActionEvent event) {
                try {
                    xls = new FileInputStream(t.getText());
                    XSSFWorkbook wb = new XSSFWorkbook(xls);
                    int sheetNo = wb.getNumberOfSheets();
                    String error = "";
                    for (int i = 0; i < sheetNo; i++) {
                        XSSFSheet sheet = wb.getSheetAt(i);
                        error += done(names, sheet, selectedDirectory);
                    }

                    don.setText("if Error :\n" + error + "Done, check your files now");
                } catch (IOException ex) {
                    Logger.getLogger(ExcelCon.class.getName()).log(Level.SEVERE, null, ex);
                }

            }

        });

        gridpane.add(into, 0, 0, 2, 1);
        gridpane.add(t, 0, 1);
        gridpane.add(btn, 1, 1);
        gridpane.add(saveP, 0, 2);
        gridpane.add(saveB, 1, 2);
        gridpane.add(run, 0, 3, 2, 1);
        gridpane.add(don, 0, 4, 2, 1);

        t.setVisible(true);
        StackPane root = new StackPane();
        ScrollPane sp = new ScrollPane();
        sp.setContent(gridpane);
        root.getChildren().add(sp);
        root.setBlendMode(BlendMode.MULTIPLY);

        Scene scene = new Scene(root, 450, 250, Color.ALICEBLUE);

        primaryStage.setTitle("Excel txt export");
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        launch(args);
    }

    public static String done(String[][] names, XSSFSheet sheet, File selectedDirectory) throws IOException {

        String headBlock = "";
        String s = "";
        int namesLength = names.length;
        String[] statment = new String[namesLength];
        File[] file = new File[names.length];
        BufferedWriter[] writer = new BufferedWriter[names.length];
        for (int i = 0; i < names.length; i++) {
            file[i] = new File(selectedDirectory.getPath()
                    + "\\" + names[i][0] + " for sheet " + sheet.getSheetName() + ".txt");

            writer[i] = new BufferedWriter(new FileWriter(file[i]));
            statment[i] = "";

        }

        for (Row row : sheet) {
            int cellIndex = 0;
            String _tempHeadBlock = row.getCell(cellIndex, Row.CREATE_NULL_AS_BLANK).toString();
            if (!_tempHeadBlock.isEmpty()) {
                headBlock = _tempHeadBlock;
            }
            if (headBlock == "") {
                System.out.println("Head Block is empty");
            }
            cellIndex++;
            //System.out.println(headBlock);
            String _tempFeature = row.getCell(cellIndex, Row.CREATE_NULL_AS_BLANK).toString();

            for (int i = 0; i < namesLength; i++) {
                if (_tempFeature.equals(names[i][0])) {
                    //System.out.println(_tempFeature);
                    cellIndex++;
                    try {
                        int no1 = Math.round(
                                Float.valueOf(
                                        row.getCell(cellIndex, Row.CREATE_NULL_AS_BLANK).toString()));
                        // System.out.println("No1 :"+ no1);
                        cellIndex++;

                        int no2 = Math.round(
                                Float.valueOf(
                                        row.getCell(cellIndex, Row.CREATE_NULL_AS_BLANK).toString()));

                        //System.out.println("No2 :"+ no2);
                        if (no1 < no2) {
                            int tempNo = no1;
                            no1 = no2;
                            no2 = tempNo;
                        }
                        statment[i] = "";
                        statment[i] += names[i][1];
                        statment[i] += no1 + ":" + no1;
                        statment[i] += "\n";
                        statment[i] += headBlock + ":" + no1
                                + ":" + no2 + "\n";
                        //System.out.println(statment[i]);
                    } catch (Exception e) {
                        int rowin = row.getRowNum();
                        rowin += 1;
                        int celle = cellIndex + 1;
                        s += "Sheet: " + sheet.getSheetName() + " Row :" + rowin + " Cell:" + celle + "\n";
                        System.out.println(sheet.getSheetName() + "Row :" + rowin + " Cell:" + celle);

                    }
                    writer[i].write(statment[i]);

                }
            }

        }
        for (int i = 0; i < names.length; i++) {
            writer[i].write("End of MpartPos list");
            writer[i].close();

        }
        return s;
    }
}
