package sample;

import javax.swing.*;


import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.stage.Popup;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;

import java.awt.*;
import java.io.*;
import java.math.BigDecimal;
import java.math.MathContext;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;


public class Controller {
    @FXML
    private TextField bottleDirTextField;
    @FXML
    private Button bottleDirButton;
    @FXML
    private TextField nox500TextField;
    @FXML
    private TextField nox2500TextField;
    @FXML
    private TextField nox10kTextField;
    @FXML
    private TextField no500TextField;
    @FXML
    private TextField no2500TextField;
    @FXML
    private TextField no10kTextField;
    @FXML
    private TextField thc500TextField;
    @FXML
    private TextField thc2500TextField;
    @FXML
    private TextField thc10kTextField;
    @FXML
    private TextField ch4500TextField;
    @FXML
    private TextField ch42500TextField;
    @FXML
    private TextField ch410kTextField;
    @FXML
    private TextField co5000TextField;
    @FXML
    private TextField cohighTextField;
    @FXML
    private TextField co216TextField;
    @FXML
    private TextField eco2TextField;
    @FXML
    private TextField o225TextField;
    @FXML
    private TextField n21TextField;
    @FXML
    private TextField n22TextField;
    @FXML
    private TextField n23TextField;
    @FXML
    private TextField n24TextField;
    @FXML
    private TextField air1TextField;
    @FXML
    private TextField air2TextField;
    @FXML
    private TextField air3TextField;
    @FXML
    private TextField air4TextField;
    @FXML
    private TextField o21001TextField;
    @FXML
    private TextField o21002TextField;
    @FXML
    private TextField o21003TextField;
    @FXML
    private TextField o21004TextField;
    @FXML
    private TextField fuel1TextField;
    @FXML
    private TextField fuel2TextField;
    @FXML
    private TextField excel1040TextField;
    @FXML
    private TextField excel1016TextField;
    @FXML
    private TextField calSheetTextField;

    //Span array to collect the information to write to excel
    ArrayList<String> spanLotList = new ArrayList<String>();
    ArrayList<String> spanCertDateList = new ArrayList<String>();
    ArrayList<String> spanExpDateList = new ArrayList<String>();
    ArrayList<String> spanValueList = new ArrayList<String>();
    ArrayList<String> spanCylNumList = new ArrayList<String>();
    ArrayList<String> spanUnitList = new ArrayList<String>();
    ArrayList<String> spanNistList = new ArrayList<String>();
    ArrayList<String> nox500List = new ArrayList<String>();
    ArrayList<String> nox2500List = new ArrayList<String>();
    ArrayList<String> nox10kList = new ArrayList<String>();
    ArrayList<String> no500List = new ArrayList<String>();
    ArrayList<String> no2500List = new ArrayList<String>();
    ArrayList<String> no10kList = new ArrayList<String>();
    ArrayList<String> thc500List = new ArrayList<String>();
    ArrayList<String> thc2500List = new ArrayList<String>();
    ArrayList<String> thc10kList = new ArrayList<String>();
    ArrayList<String> ch4500List = new ArrayList<String>();
    ArrayList<String> ch42500List = new ArrayList<String>();
    ArrayList<String> ch410kList = new ArrayList<String>();
    ArrayList<String> co5000List = new ArrayList<String>();
    ArrayList<String> coHighList = new ArrayList<String>();
    ArrayList<String> co216List = new ArrayList<String>();
    ArrayList<String> egr16List = new ArrayList<String>();
    ArrayList<String> o225List = new ArrayList<String>();
    ArrayList<String> n21List = new ArrayList<String>();
    ArrayList<String> n22List = new ArrayList<String>();
    ArrayList<String> n23List = new ArrayList<String>();
    ArrayList<String> n24List = new ArrayList<String>();
    ArrayList<String> air1List = new ArrayList<String>();
    ArrayList<String> air2List = new ArrayList<String>();
    ArrayList<String> air3List = new ArrayList<String>();
    ArrayList<String> air4List = new ArrayList<String>();
    ArrayList<String> o21001List = new ArrayList<String>();
    ArrayList<String> o21002List = new ArrayList<String>();
    ArrayList<String> o21003List = new ArrayList<String>();
    ArrayList<String> o21004List = new ArrayList<String>();
    ArrayList<String> fuel1List = new ArrayList<String>();
    ArrayList<String> fuel2List = new ArrayList<String>();
    ArrayList<String> cylNumVerificationList = new ArrayList<String>();

    PDDocument fileDocument;


    public Controller() throws IOException {
    }


    public void setBottleDirectory() {


        // Set bottle directory

        String userDir = System.getProperty("user.home");

        JFileChooser j = new JFileChooser(userDir + "/Desktop");

        j.setDialogTitle("Select Bottle Folder");

        j.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

        j.setAcceptAllFileFilterUsed(false);

        j.showOpenDialog(null);


        bottleDirTextField.setText(j.getCurrentDirectory().getPath() + "\\" + "Bottles");

        String horibaDirectory = bottleDirTextField.getText().replace("\\Bottles", "");

        String bottleDirectory = bottleDirTextField.getText();


        File folder = new File(bottleDirectory);

        File[] listOfFiles = folder.listFiles();


        for (int i = 0; i < listOfFiles.length; i++) {

            if (listOfFiles[i].toString().contains("Nox 500"))
                nox500TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("Nox 2500"))
                nox2500TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("Nox 10000"))
                nox10kTextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("THC 500"))
                thc500TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("THC 2500"))
                thc2500TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("THC 10000"))
                thc10kTextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("CH4 500"))
                ch4500TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("CH4 2500"))
                ch42500TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("CH4 10000"))
                ch410kTextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("CO(L) 5000"))
                co5000TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("CO(H) 1.6%"))
                cohighTextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("CO2 16%"))
                co216TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("EGR 16%"))
                eco2TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("O2 25%"))
                o225TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("N2 1"))
                n21TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("N2 2"))
                n22TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("N2 3"))
                n23TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("N2 4"))
                n24TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("Air 1"))
                air1TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("Air 2"))
                air2TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("Air 3"))
                air3TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("Air 4"))
                air4TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("Fuel 1"))
                fuel1TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("Fuel 2"))
                fuel2TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("NO(C) 500"))
                no500TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("NO(C) 2500"))
                no2500TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("NO(C) 10000"))
                no10kTextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("O2 100% 1"))
                o21001TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("O2 100% 2"))
                o21002TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("O2 100% 3"))
                o21003TextField.setText(listOfFiles[i].getName());
            if (listOfFiles[i].toString().contains("O2 100% 4"))
                o21004TextField.setText(listOfFiles[i].getName());
        }

        File horibafolder = new File(horibaDirectory);
        File[] listOfTemplates = horibafolder.listFiles();

        for (int x = 0; x < listOfTemplates.length; x++) {
            if (listOfTemplates[x].toString().contains("Horiba_Cal_Sheet"))
                calSheetTextField.setText(listOfTemplates[x].getName().replace("~$",""));
        }

    }

    public void setCalSheetDirectory() {

        // Set bottle directory

        String userDir = System.getProperty("user.home");

        JFileChooser j = new JFileChooser(userDir + "/Desktop");

        j.setDialogTitle("Select Cal Sheet");

        j.setFileSelectionMode(JFileChooser.FILES_ONLY);

        j.setAcceptAllFileFilterUsed(false);

        j.showOpenDialog(null);


        calSheetTextField.setText(j.getSelectedFile().getName().replace("~$",""));

    }

    public void resetForm() {
        bottleDirTextField.setText("");

        nox500TextField.setText("");

        nox2500TextField.setText("");

        nox10kTextField.setText("");

        thc500TextField.setText("");

        thc2500TextField.setText("");

        thc10kTextField.setText("");

        ch4500TextField.setText("");

        ch42500TextField.setText("");

        ch410kTextField.setText("");

        co5000TextField.setText("");

        cohighTextField.setText("");

        co216TextField.setText("");

        eco2TextField.setText("");

        o225TextField.setText("");

        n21TextField.setText("");

        n22TextField.setText("");

        n23TextField.setText("");

        n24TextField.setText("");

        air1TextField.setText("");

        air2TextField.setText("");

        air3TextField.setText("");

        air4TextField.setText("");

        fuel1TextField.setText("");

        fuel2TextField.setText("");

        no500TextField.setText("");

        no2500TextField.setText("");

        no10kTextField.setText("");

        o21001TextField.setText("");

        o21002TextField.setText("");

        o21003TextField.setText("");

        o21004TextField.setText("");

        calSheetTextField.setText("");
    }


    public String stripPDF(String fileName) throws ParseException, IOException {


        File file = new File(fileName);

        //Opening 1016

        File excel1016File = new File(bottleDirTextField.getText().replace("Bottles", "") + calSheetTextField.getText());


        FileInputStream excel1016InputStream = new FileInputStream(excel1016File);

        Workbook excel1016Workbook = WorkbookFactory.create(excel1016InputStream);

        Sheet excel1016Sheet = excel1016Workbook.getSheetAt(0);


        //Instantiate PDFTextStripper class

        PDFTextStripper pdfStripper = null;


        fileDocument = PDDocument.load(file);


        pdfStripper = new PDFTextStripper();


        //Retrieving text from PDF document

        return pdfStripper.getText(fileDocument);

    }


    public void nox500Extraction() throws IOException, ParseException {


        if (!"".equals(nox500TextField.getText())) {

            String nox500cylNum = null;

            File file = new File(bottleDirTextField.getText() + "\\" + nox500TextField.getText());

            String fileText = stripPDF(bottleDirTextField.getText() + "\\" + nox500TextField.getText());

            String[] fileTextArray = fileText.split("\\r\\n");

            for (int i = 0; i < fileTextArray.length; i++) {


                if (fileTextArray[i].contains("Lot Number:")) {
                    //Lot number
                    String nox500Lot = fileTextArray[i].replace("Lot Number:", "");

                    nox500List.add(nox500Lot);


                }
                if (fileTextArray[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String nox500certDate = fileTextArray[i - 1];

                    SimpleDateFormat nox500formatter = new SimpleDateFormat("M/d/yyyy");
                    Date nox500FormattedDate = nox500formatter.parse(nox500certDate);
                    nox500List.add(new SimpleDateFormat("d-MMM-yyyy").format(nox500FormattedDate));


                }
                if (fileTextArray[i].contains("Primary Master")) {

                    //Expiration Date
                    String nox500expDate = fileTextArray[i + 1].replace("Expiration Date: ", "");

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(nox500expDate);
                    nox500List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));


                }
                if (fileTextArray[i].contains("For Reference Only:")) {

                    //Nox 500 Span value
                    String nox500span = fileTextArray[i].substring(24, 27);

                    nox500List.add(nox500span);


                }

                //Cylinder number from file name
                String nox500cyl = file.getName();
                //Strip out the gas type from file name
                String nox500blankText1 = nox500cyl.replace("Nox 500 ", "");
                //Strip out .pdf from file name
                String nox500blankText2 = nox500blankText1.replace(".pdf", "");
                //Actual Cylinder number from file name(Must be Nox 10k cyl#)
                nox500cylNum = nox500blankText2;

                if (fileTextArray[i].contains(nox500cylNum)) {

                    nox500List.add(nox500cylNum);


                }
            }

            if (nox500List.toString().contains(nox500cylNum)) {

                cylNumVerificationList.add("Nox 500 Cylinder Number Verified");
            } else {

                cylNumVerificationList.add("Nox 500 Error");
                nox500List.add("ERROR");
            }

           

            nox500List.add("ppm");
            nox500List.add("1% NIST");


            fileDocument.close();

        }
    }

    public void nox2500extraction() throws IOException, ParseException {

        if (!"".equals(nox2500TextField.getText())) {

            String nox2500cylNum = null;

            File file = new File(bottleDirTextField.getText() + "\\" + nox2500TextField.getText());

            String fileText = stripPDF(bottleDirTextField.getText() + "\\" + nox2500TextField.getText());

            String[] fileTextArray = fileText.split("\\r\\n");

            for (int i = 0; i < fileTextArray.length; i++) {


                if (fileTextArray[i].contains("Lot Number:")) {

                    //Lot number

                    String nox2500Lot = fileTextArray[i].replace("Lot Number:", "");


                    nox2500List.add(nox2500Lot);


                }
                if (fileTextArray[i].contains("Certificate Issuance Date:")) {


                    //Cert Date

                    String nox2500certDate = fileTextArray[i - 1];


                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");

                    Date newFormattedDate = dateformatter.parse(nox2500certDate);

                    nox2500List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));


                }
                if (fileTextArray[i].contains("Primary Master")) {


                    //Expiration Date

                    String nox2500expDate = fileTextArray[i + 1].replace("Expiration Date: ", "");


                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");

                    Date newFormattedDate = dateformatter.parse(nox2500expDate);

                    nox2500List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));


                }
                if (fileTextArray[i].contains("For Reference Only:")) {


                    //nox 2500 2500 Span value

                    String nox2500span = fileTextArray[i].substring(24, 28);


                    nox2500List.add(nox2500span);


                }

                //Cylinder number from file name

                String nox2500cyl = file.getName();

                //Strip out the gas type from file name

                String nox2500blankText1 = nox2500cyl.replace("Nox 2500 ", "");

                //Strip out .pdf from file name

                String nox2500blankText2 = nox2500blankText1.replace(".pdf", "");

                //Actual Cylinder number from file name(Must be Nox 2500 cyl#)

                nox2500cylNum = nox2500blankText2;

                if (fileTextArray[i].contains(nox2500cylNum)) {

                    nox2500List.add(nox2500cylNum);

                }
            }


            if (nox2500List.toString().contains(nox2500cylNum)) {

                cylNumVerificationList.add("Nox 2500 Cylinder Number Verified");
            } else {

                cylNumVerificationList.add("Nox 2500 Error");
                nox2500List.add("ERROR");
            }

  
            nox2500List.add("ppm");
            nox2500List.add("1% NIST");


            fileDocument.close();

        }
    }


    public void nox10kExtraction() throws IOException, ParseException {


        if (!"".equals(nox10kTextField.getText())) {

            String nox10kCylNum = null;

            File file = new File(bottleDirTextField.getText() + "\\" + nox10kTextField.getText());

            String fileText = stripPDF(bottleDirTextField.getText() + "\\" + nox10kTextField.getText());

            String[] fileTextArray = fileText.split("\\r\\n");

            for (int i = 0; i < fileTextArray.length; i++) {

                if (fileTextArray[i].contains("Lot Number:")) {

                    //Lot number

                    String nox10kLot = fileTextArray[i].replace("Lot Number:", "");


                    nox10kList.add(nox10kLot);


                }
                if (fileTextArray[i].contains("Certificate Issuance Date:")) {

                    //Cert Date

                    String nox10kcertDate = fileTextArray[i - 1];


                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");

                    Date newFormattedDate = dateformatter.parse(nox10kcertDate);

                    nox10kList.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));


                }
                if (fileTextArray[i].contains("Primary Master")) {

                    //Expiration Date

                    String nox10kexpDate = fileTextArray[i + 1].replace("Expiration Date: ", "");


                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");

                    Date newFormattedDate = dateformatter.parse(nox10kexpDate);

                    nox10kList.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));


                }
                if (fileTextArray[i].contains("For Reference Only:")) {

                    //Nox 10k Span value

                    String nox10kspan = fileTextArray[i].substring(24, 28);

                    nox10kList.add(nox10kspan);


                }

                //Cylinder number from file name

                String nox10kcyl = file.getName();

                //Strip out the gas type from file name

                String nox10kblankText1 = nox10kcyl.replace("Nox 10000 ", "");

                //Strip out .pdf from file name

                String nox10kblankText2 = nox10kblankText1.replace(".pdf", "");

                //Actual Cylinder number from file name(Must be Nox 10k cyl#)

                nox10kCylNum = nox10kblankText2;
                if (fileTextArray[i].contains(nox10kCylNum)) {


                    nox10kList.add(nox10kCylNum);

                }
            }

            if (nox10kList.toString().contains(nox10kCylNum)) {
                cylNumVerificationList.add("Nox 10k Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("Nox 10k Cylinder Number Error");
            }

            if(nox10kList.size() < 5 ) {

                nox10kList.add(1, "ERROR");
            }

            nox10kList.add("ppm");

            nox10kList.add("1% NIST");


        }
    }

    public void no500Extraction() throws IOException, ParseException {

        if (!"".equals(no500TextField.getText())) {

            String no500cylNum = null;

            File file = new File(bottleDirTextField.getText() + "\\" + no500TextField.getText());

            String fileText = stripPDF(bottleDirTextField.getText() + "\\" + no500TextField.getText());

            String[] fileTextArray = fileText.split("\\r\\n");

            for (int i = 0; i < fileTextArray.length; i++) {

                if (fileTextArray[i].contains("Lot Number:")) {
                    //Lot number
                    String no500Lot = fileTextArray[i].replace("Lot Number:", "");

                    no500List.add(no500Lot);

                }
                if (fileTextArray[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String no500certDate = fileTextArray[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(no500certDate);
                    no500List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Primary Master")) {

                    //Expiration Date
                    String no500expDate = fileTextArray[i + 1].replace("Expiration Date: ", "");

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(no500expDate);
                    no500List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Nitric oxide")) {

                    //NO(C) 500 Span value
                    String no500replaceText1 = fileTextArray[i].replace("Nitric oxide 2.5 Chemically Pure (CP) 475 ppm ", "");
                    String no500replaceText2 = no500replaceText1.replace("Nitric oxide 475 ppm ", "");
                    String no500span = no500replaceText2.replace(" ppm 1", "");

                    no500List.add(no500span);

                }

                //Cylinder number from file name
                String no500cyl = file.getName();
                //Strip out the gas type from file name
                String no500blankText1 = no500cyl.replace("NO(C) 500 ", "");
                //Strip out .pdf from file name
                String no500blankText2 = no500blankText1.replace(".pdf", "");
                //Actual Cylinder number from file name(Must be Nox 2500 cyl#)
                no500cylNum = no500blankText2;

                if (fileTextArray[i].contains(no500cylNum)) {
                    no500List.add(no500cylNum);
                }

            }

            if (nox500List.toString().contains(no500cylNum)) {
                cylNumVerificationList.add("NO(C) 500 Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("NO(C) 500 Cylinder Number Error");
            }

            if(no500List.size() < 5 ) {

                no500List.add(0, "ERROR");
            }



            no500List.add("ppm");
            no500List.add("1% NIST");


        }
    }

    public void no2500Extraction() throws IOException, ParseException {

        if (!"".equals(no2500TextField.getText())) {

            String no2500cylNum = null;

            File file = new File(bottleDirTextField.getText() + "\\" + no2500TextField.getText());

            String fileText = stripPDF(bottleDirTextField.getText() + "\\" + no2500TextField.getText());

            String[] fileTextArray = fileText.split("\\r\\n");

            for (int i = 0; i < fileTextArray.length; i++) {

                if (fileTextArray[i].contains("Lot Number:")) {
                    //Lot number
                    String no2500Lot = fileTextArray[i].replace("Lot Number:", "");

                    no2500List.add(no2500Lot);

                }
                if (fileTextArray[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String no2500certDate = fileTextArray[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(no2500certDate);
                    no2500List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Primary Master")) {

                    //Expiration Date
                    String no2500expDate = fileTextArray[i + 1].replace("Expiration Date: ", "");

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(no2500expDate);
                    no2500List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Nitric oxide")) {

                    //NO(C) 2500 Span value
                    String no2500replaceText1 = fileTextArray[i].replace("Nitric oxide 2.5 Chemically Pure (CP) 2375 ppm ", "");
                    String no2500replaceText2 = no2500replaceText1.replace("Nitric oxide 2375 ppm ", "");
                    String no2500span = no2500replaceText2.replace(" ppm 1", "");

                    no2500List.add(no2500span);

                }

                //Cylinder number from file name
                String no2500cyl = file.getName();
                //Strip out the gas type from file name
                String no2500blankText1 = no2500cyl.replace("NO(C) 2500 ", "");
                //Strip out .pdf from file name
                String no2500blankText2 = no2500blankText1.replace(".pdf", "");
                //Actual Cylinder number from file name(Must be Nox 2500 cyl#)
                no2500cylNum = no2500blankText2;

                if (fileTextArray[i].contains(no2500cylNum)) {
                    no2500List.add(no2500cylNum);
                }

            }

            if (no2500List.toString().contains(no2500cylNum)) {
                cylNumVerificationList.add("NO(C) 2500 Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("NO(C) 2500 Cylinder Number Error");
            }

            if(no2500List.size() < 5 ) {

                no2500List.add(0, "ERROR");
            }



            no2500List.add("ppm");
            no2500List.add("1% NIST");


        }
    }

    public void no10kExtraction() throws IOException, ParseException {

        if (!"".equals(no10kTextField.getText())) {

            String no10kcylNum = null;

            File file = new File(bottleDirTextField.getText() + "\\" + no10kTextField.getText());

            String fileText = stripPDF(bottleDirTextField.getText() + "\\" + no10kTextField.getText());

            String[] fileTextArray = fileText.split("\\r\\n");

            for (int i = 0; i < fileTextArray.length; i++) {

                if (fileTextArray[i].contains("Lot Number:")) {
                    //Lot number
                    String no10kLot = fileTextArray[i].replace("Lot Number:", "");

                    no10kList.add(no10kLot);

                }
                if (fileTextArray[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String no10kcertDate = fileTextArray[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(no10kcertDate);
                    no10kList.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Primary Master")) {

                    //Expiration Date
                    String no10kexpDate = fileTextArray[i + 1].replace("Expiration Date: ", "");

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(no10kexpDate);
                    no10kList.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Nitric oxide")) {

                    //Nox 10k Span value
                    String no10kreplaceText1 = fileTextArray[i].replace("Nitric oxide 9500 ppm ", "");
                    String no10kreplaceText2 = no10kreplaceText1.replace("Nitric oxide 2.5 Chemically Pure (CP) 9500 ppm ", "");
                    String no10kspan = no10kreplaceText2.substring(0, 4);

                    no10kList.add(no10kspan);

                }

                //Cylinder number from file name
                String no10kcyl = file.getName();
                //Strip out the gas type from file name
                String no10kblankText1 = no10kcyl.replace("NO(C) 10000 ", "");
                //Strip out .pdf from file name
                String no10kblankText2 = no10kblankText1.replace(".pdf", "");
                //Actual Cylinder number from file name(Must be Nox 10k cyl#)
                no10kcylNum = no10kblankText2;

                if (fileTextArray[i].contains(no10kcylNum)) {


                    no10kList.add(no10kcylNum);
                }
            }
            if (no10kList.toString().contains(no10kcylNum)) {
                cylNumVerificationList.add("NO(C) 10K Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("NO(C) 10K Cylinder Number Error");
            }

            if(no10kList.size() < 5 ) {

                no10kList.add(0, "ERROR");
            }

            no10kList.add("ppm");
            no10kList.add("1% NIST");


        }

    }

    public void thc500Extraction() throws IOException, ParseException {

        if (!"".equals(thc500TextField.getText())) {

            String thc500cylNum = null;

            File file = new File(bottleDirTextField.getText() + "\\" + thc500TextField.getText());

            String fileText = stripPDF(bottleDirTextField.getText() + "\\" + thc500TextField.getText());

            String[] fileTextArray = fileText.split("\\r\\n");

            for (int i = 0; i < fileTextArray.length; i++) {

                if (fileTextArray[i].contains("Lot Number:")) {
                    //Lot number
                    String thc500Lot = fileTextArray[i].replace("Lot Number:", "");

                    thc500List.add(thc500Lot);

                }
                if (fileTextArray[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String thc500certDate = fileTextArray[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(thc500certDate);
                    thc500List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Primary Master")) {

                    //Expiration Date
                    String thc500expDate = fileTextArray[i + 1].replace("Expiration Date: ", "");

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(thc500expDate);
                    thc500List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Propane ")) {

                    //THC 500 Span value
                    String thc500replaceText1 = fileTextArray[i].replace("Propane 158 ppm ", "");
                    String thc500replaceText2 = thc500replaceText1.replace("Propane 4.0 Research 158 ppm ", "");
                    String thc500replaceText3 = thc500replaceText2.replace("Propane 2.5 Instrument 158 ppm ", "");
                    String thc500span = thc500replaceText3.replace(" ppm 1", "");
                    double thc500newspan = Double.parseDouble(thc500span);
                    double thc500spanx3 = thc500newspan * 3;
                    BigDecimal thc500spanDecimal = new BigDecimal(Double.toString(thc500spanx3));
                    MathContext mc = new MathContext(4);
                    BigDecimal thc500spanRounded = thc500spanDecimal.round(mc);
                    String thc500spanConverted = thc500spanRounded.toString();


                    thc500List.add(thc500spanConverted);

                }

                //Cylinder number from file name
                String thc500cyl = file.getName();
                //Strip out the gas type from file name
                String thc500blankText1 = thc500cyl.replace("THC 500 ", "");
                //Strip out .pdf from file name
                String thc500blankText2 = thc500blankText1.replace(".pdf", "");
                //Actual Cylinder number from file name(Must be Fuel cyl#)
                thc500cylNum = thc500blankText2;

                if (fileTextArray[i].contains(thc500cylNum)) {

                    thc500List.add(thc500cylNum);


                }
            }

            if (thc500List.toString().contains(thc500cylNum)) {
                cylNumVerificationList.add("THC 500 Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("THC 500 Cylinder Number Error");
            }

            if(thc500List.size() < 5 ) {

                thc500List.add(0, "ERROR");
            }


            thc500List.add("ppm");
            thc500List.add("1% NIST");


        }

    }

    public void thc2500Extraction() throws IOException, ParseException {

        if (!"".equals(thc2500TextField.getText())) {

            String thc2500cylNum = null;

            File file = new File(bottleDirTextField.getText() + "\\" + thc2500TextField.getText());

            String fileText = stripPDF(bottleDirTextField.getText() + "\\" + thc2500TextField.getText());

            String[] fileTextArray = fileText.split("\\r\\n");

            for (int i = 0; i < fileTextArray.length; i++) {


                if (fileTextArray[i].contains("Lot Number:")) {
                    //Lot number
                    String thc2500Lot = fileTextArray[i].replace("Lot Number:", "");

                    thc2500List.add(thc2500Lot);

                }
                if (fileTextArray[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String thc2500certDate = fileTextArray[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(thc2500certDate);
                    thc2500List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Primary Master")) {

                    //Expiration Date
                    String thc2500expDate = fileTextArray[i + 1].replace("Expiration Date: ", "");

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(thc2500expDate);
                    thc2500List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Propane")) {

                    //THC 2500 Span value
                    String thc2500replaceText1 = fileTextArray[i].replace("Propane 2.5 Instrument 792 ppm ", "");
                    String thc2500replaceText2 = thc2500replaceText1.replace("Propane 792 ppm ", "");
                    String thc2500replaceText3 = thc2500replaceText2.replace("Propane 4.0 Research 792 ppm ", "");
                    String thc2500span = thc2500replaceText3.replace(" ppm 1", "");
                    int thc2500spanint = Integer.parseInt(thc2500span);
                    int thc2500spanx3 = thc2500spanint * 3;
                    String thc2500newspan = String.valueOf(thc2500spanx3);

                    thc2500List.add(thc2500newspan);

                }

                //Cylinder number from file name
                String thc2500cyl = file.getName();
                //Strip out the gas type from file name
                String thc2500blankText1 = thc2500cyl.replace("THC 2500 ", "");
                //Strip out .pdf from file name
                String thc2500blankText2 = thc2500blankText1.replace(".pdf", "");
                //Actual Cylinder number from file name(Must be Fuel cyl#)
                thc2500cylNum = thc2500blankText2;

                if (fileTextArray[i].contains(thc2500cylNum)) {

                    thc2500List.add(thc2500cylNum);

                }
            }

            if (thc2500List.toString().contains(thc2500cylNum)) {

                cylNumVerificationList.add("THC 2500 Cylinder Number Verified");

            } else {

                cylNumVerificationList.add("THC 2500 Cylinder Number Error");
            }

            if(thc2500List.size() < 5 ) {

                thc2500List.add(0, "ERROR");
            }

            thc2500List.add("ppm");
            thc2500List.add("1% NIST");


        }
    }

    public void thc10kExtraction() throws IOException, ParseException {

        if (!"".equals(thc10kTextField.getText())) {

            String thc10kcylNum = null;

            File file = new File(bottleDirTextField.getText() + "\\" + thc10kTextField.getText());

            String fileText = stripPDF(bottleDirTextField.getText() + "\\" + thc10kTextField.getText());

            String[] fileTextArray = fileText.split("\\r\\n");

            for (int i = 0; i < fileTextArray.length; i++) {

                if (fileTextArray[i].contains("Lot Number:")) {
                    //Lot number
                    String thc10kLot = fileTextArray[i].replace("Lot Number:", "");

                    thc10kList.add(thc10kLot);

                }
                if (fileTextArray[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String thc10kcertDate = fileTextArray[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(thc10kcertDate);
                    thc10kList.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Primary Master")) {

                    //Expiration Date
                    String thc10kexpDate = fileTextArray[i + 1].replace("Expiration Date: ", "");

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(thc10kexpDate);

                    thc10kList.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Propane")) {

                    //THC 10k Span value
                    String thc10kreplaceText1 = fileTextArray[i].replace("Propane 2.5 Instrument 3200 ppm ", "");
                    String thc10kreplaceText2 = thc10kreplaceText1.replace("Propane 3200 ppm ", "");
                    String thc10kreplaceText3 = thc10kreplaceText2.replace("Propane 4.0 Research 3200 ppm ", "");
                    String thc10kspan = thc10kreplaceText3.substring(0, 4);
                    int thc10kspanint = Integer.parseInt(thc10kspan);
                    int thc10kspanx3 = thc10kspanint * 3;
                    String thc10knewspan = String.valueOf(thc10kspanx3);

                    thc10kList.add(thc10knewspan);

                }

                //Cylinder number from file name
                String thc10kcyl = file.getName();
                //Strip out the gas type from file name
                String thc10kblankText1 = thc10kcyl.replace("THC 10000 ", "");
                //Strip out .pdf from file name
                String thc10kblankText2 = thc10kblankText1.replace(".pdf", "");
                //Actual Cylinder number from file name(Must be THC 10k cyl#)
                thc10kcylNum = thc10kblankText2;

                if (fileTextArray[i].contains(thc10kcylNum)) {

                    thc10kList.add(thc10kcylNum);

                }
            }

            if (thc10kList.toString().contains(thc10kcylNum)) {

                cylNumVerificationList.add("THC 10K Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("THC 10K Cylinder Number Error");
            }

            if(thc10kList.size() < 5 ) {

                thc10kList.add(0, "ERROR");
            }


            thc10kList.add("ppm");
            thc10kList.add("1% NIST");


        }

    }

    public void ch4500Extraction() throws IOException, ParseException {

        if (!"".equals(ch4500TextField.getText())) {

            String ch4500cylNum = null;


            File file = new File(bottleDirTextField.getText() + "\\" + ch4500TextField.getText());

            String fileText = stripPDF(bottleDirTextField.getText() + "\\" + ch4500TextField.getText());

            String[] fileTextArray = fileText.split("\\r\\n");

            for (int i = 0; i < fileTextArray.length; i++) {

                if (fileTextArray[i].contains("Lot Number:")) {
                    //Lot number
                    String ch4500Lot = fileTextArray[i].replace("Lot Number:", "");

                    ch4500List.add(ch4500Lot);

                }
                if (fileTextArray[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String ch4500certDate = fileTextArray[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(ch4500certDate);
                    ch4500List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Primary Master")) {

                    //Expiration Date
                    String ch4500expDate = fileTextArray[i + 1].replace("Expiration Date: ", "");

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(ch4500expDate);
                    ch4500List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Methane")) {

                    //CH4 500 Span value
                    String ch4500replaceText1 = fileTextArray[i].replace("Methane 3.7 Ultra High Purity 475 ppm ", "");
                    String ch4500replaceText2 = ch4500replaceText1.replace("Methane 475 ppm ", "");
                    String ch4500span = ch4500replaceText2.replace(" ppm 1", "");

                    ch4500List.add(ch4500span);

                }

                //Cylinder number from file name
                String ch4500cyl = file.getName();
                //Strip out the gas type from file name
                String ch4500blankText1 = ch4500cyl.replace("CH4 500 ", "");
                //Strip out .pdf from file name
                String ch4500blankText2 = ch4500blankText1.replace(".pdf", "");
                //Actual Cylinder number from file name(Must be Fuel cyl#)

                ch4500cylNum = ch4500blankText2;

                if (fileTextArray[i].contains(ch4500cylNum)) {
                    ch4500List.add(ch4500cylNum);
                }
            }

            if (ch4500List.toString().contains(ch4500cylNum)) {

                cylNumVerificationList.add("CH4 500 Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("Ch4 500 Cylinder Number Error");
            }

            if(ch4500List.size() < 5 ) {

                ch4500List.add(0, "ERROR");
            }


            ch4500List.add("ppm");
            ch4500List.add("1% NIST");


        }

    }


    public void ch42500Extraction() throws IOException, ParseException {
        if (!"".equals(ch42500TextField.getText())) {

            String ch42500cylNum = null;

            File file = new File(bottleDirTextField.getText() + "\\" + ch42500TextField.getText());

            String fileText = stripPDF(bottleDirTextField.getText() + "\\" + ch42500TextField.getText());

            String[] fileTextArray = fileText.split("\\r\\n");

            for (int i = 0; i < fileTextArray.length; i++) {

                if (fileTextArray[i].contains("Lot Number:")) {
                    //Lot number
                    String ch42500Lot = fileTextArray[i].replace("Lot Number:", "");

                    ch42500List.add(ch42500Lot);

                }
                if (fileTextArray[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String ch42500certDate = fileTextArray[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(ch42500certDate);
                    ch42500List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Primary Master")) {

                    //Expiration Date
                    String ch42500expDate = fileTextArray[i + 1].replace("Expiration Date: ", "");

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(ch42500expDate);
                    ch42500List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Methane 2500")) {

                    //CH4 2500 Span value
                    String ch42500replaceText1 = fileTextArray[i].replace("Methane 2500 ppm ", "");
                    String ch42500replaceText2 = ch42500replaceText1.replace("Methane 3.7 Ultra High Purity 2500 ppm ", "");
                    String ch42500span = ch42500replaceText2.replace(" ppm 1", "");

                    ch42500List.add(ch42500span);

                }

                //Cylinder number from file name
                String ch42500cyl = file.getName();
                //Strip out the gas type from file name
                String ch42500blankText1 = ch42500cyl.replace("CH4 2500 ", "");
                //Strip out .pdf from file name
                String ch42500blankText2 = ch42500blankText1.replace(".pdf", "");
                //Actual Cylinder number from file name(Must be Fuel cyl#)
                ch42500cylNum = ch42500blankText2;

                if (fileTextArray[i].contains(ch42500cylNum)) {

                    ch42500List.add(ch42500cylNum);

                }
            }

            if (ch42500List.toString().contains(ch42500cylNum)) {

                cylNumVerificationList.add("CH4 2500 Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("CH4 2500 Cylinder Number Error");
            }

            if(ch42500List.size() < 5 ) {

                ch42500List.add(0, "ERROR");
            }



            ch42500List.add("ppm");
            ch42500List.add("1% NIST");

        }

    }

    public void ch410kExtraction() throws ParseException, IOException {

        if (!"".equals(ch410kTextField.getText())) {

            String ch410kcylNum = null;

            File file = new File(bottleDirTextField.getText() + "\\" + ch410kTextField.getText());

            String fileText = stripPDF(bottleDirTextField.getText() + "\\" + ch410kTextField.getText());

            String[] fileTextArray = fileText.split("\\r\\n");

            for (int i = 0; i < fileTextArray.length; i++) {

                if (fileTextArray[i].contains("Lot Number:")) {
                    //Lot number
                    String ch410kLot = fileTextArray[i].replace("Lot Number:", "");

                    ch410kList.add(ch410kLot);

                }
                if (fileTextArray[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String ch410kcertDate = fileTextArray[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(ch410kcertDate);
                    ch410kList.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Primary Master")) {

                    //Expiration Date
                    String ch410kexpDate = fileTextArray[i + 1].replace("Expiration Date: ", "");

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(ch410kexpDate);
                    ch410kList.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Methane")) {

                    //SO(L) Span value
                    String ch410kreplaceText1 = fileTextArray[i].replace("Methane 3.7 Ultra High Purity 9500 ppm ", "");
                    String ch410kreplaceText2 = ch410kreplaceText1.replace("Methane 9500 ppm ", "");
                    String ch410kspan = ch410kreplaceText1.substring(0, 4);

                    ch410kList.add(ch410kspan);


                }

                //Cylinder number from file name
                String ch410kcyl = file.getName();
                //Strip out the gas type from file name
                String ch410kblankText1 = ch410kcyl.replace("CH4 10000 ", "");
                //Strip out .pdf from file name
                String ch410kblankText2 = ch410kblankText1.replace(".pdf", "");
                //Actual Cylinder number from file name(Must be Fuel cyl#)
                ch410kcylNum = ch410kblankText2;

                if (fileTextArray[i].contains(ch410kcylNum)) {

                    ch410kList.add(ch410kcylNum);

                }
            }

            if (ch410kList.toString().contains(ch410kcylNum)) {

                cylNumVerificationList.add("CH4 10K Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("CH4 10K Cylinder Number Error");
            }

            if(ch410kList.size() < 5 ) {

                ch410kList.add(0, "ERROR");
            }


            ch410kList.add("ppm");
            ch410kList.add("1% NIST");

        }

    }

    public void co5000Extraction() throws IOException, ParseException {

        if (!"".equals(co5000TextField.getText())) {

            String co5000cylNum = null;

            File file = new File(bottleDirTextField.getText() + "\\" + co5000TextField.getText());

            String fileText = stripPDF(bottleDirTextField.getText() + "\\" + co5000TextField.getText());

            String[] fileTextArray = fileText.split("\\r\\n");

            for (int i = 0; i < fileTextArray.length; i++) {

                if (fileTextArray[i].contains("Lot Number:")) {
                    //Lot number
                    String co5000Lot = fileTextArray[i].replace("Lot Number:", "");

                    co5000List.add(co5000Lot);

                }
                if (fileTextArray[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String co5000certDate = fileTextArray[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(co5000certDate);
                    co5000List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Primary Master")) {

                    //Expiration Date
                    String co5000expDate = fileTextArray[i + 1].replace("Expiration Date: ", "");

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(co5000expDate);
                    co5000List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Carbon monoxide")) {

                    //CO(L) Span value
                    String co5000span = fileTextArray[i].substring(38, 43);

                    co5000List.add(co5000span);

                }

                //Cylinder number from file name
                String co5000cyl = file.getName();
                //Strip out the gas type from file name
                String co5000blankText1 = co5000cyl.replace("CO(L) 5000 ", "");
                //Strip out .pdf from file name
                String co5000blankText2 = co5000blankText1.replace(".pdf", "");
                //Actual Cylinder number from file name(Must be Fuel cyl#)
                co5000cylNum = co5000blankText2;

                if (fileTextArray[i].contains(co5000cylNum)) {

                    co5000List.add(co5000cylNum);


                }
            }

            if (co5000List.toString().contains(co5000cylNum)) {

                cylNumVerificationList.add("CO(L) 5000 Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("CO(L) 5000 Cylinder Number Error");
            }

            if(co5000List.size() < 5 ) {

                co5000List.add(0, "ERROR");
            }


            co5000List.add("ppm");
            co5000List.add("1% NIST");


        }

    }

    public void coHighExtraction() throws ParseException, IOException {

        if (!"".equals(cohighTextField.getText())) {

            String cohighcylNum = null;

            File file = new File(bottleDirTextField.getText() + "\\" + cohighTextField.getText());

            String fileText = stripPDF(bottleDirTextField.getText() + "\\" + cohighTextField.getText());

            String[] fileTextArray = fileText.split("\\r\\n");

            for (int i = 0; i < fileTextArray.length; i++) {

                if (fileTextArray[i].contains("Lot Number:")) {
                    //Lot number
                    String cohighLot = fileTextArray[i].replace("Lot Number:", "");

                    coHighList.add(cohighLot);

                }
                if (fileTextArray[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String cohighcertDate = fileTextArray[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(cohighcertDate);
                    coHighList.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Primary Master")) {

                    //Expiration Date
                    String cohighexpDate = fileTextArray[i + 1].replace("Expiration Date: ", "");

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(cohighexpDate);
                    coHighList.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Carbon monoxide")) {

                    //CO(H) Span value
                    String cohighspan = fileTextArray[i].substring(37, 42);
                    BigDecimal cohighspanint = new BigDecimal(cohighspan);
                    BigDecimal cohighspandecimal = new BigDecimal(10000.000);
                    String cohighspanx10000 = cohighspanint.multiply(cohighspandecimal).toString();
                    int cohighconversion = (int) Double.parseDouble(cohighspanx10000);
                    String cohighnewspan = Integer.toString(cohighconversion);
                    coHighList.add(cohighnewspan);

                }

                //Cylinder number from file name
                String cohighcyl = file.getName();
                //Strip out the gas type from file name
                String cohighblankText1 = cohighcyl.replace("CO(H) 1.6% ", "");
                //Strip out .pdf from file name
                String cohighblankText2 = cohighblankText1.replace(".pdf", "");
                //Actual Cylinder number from file name(Must be Fuel cyl#)
                cohighcylNum = cohighblankText2;

                if (fileTextArray[i].contains(cohighcylNum)) {

                    coHighList.add(cohighcylNum);
                }
            }

            if (coHighList.toString().contains(cohighcylNum)) {

                cylNumVerificationList.add("CO(H) 1.6% Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("CO(H) 1.6% Cylinder Number Error");
            }

            if(coHighList.size() < 5 ) {

                coHighList.add(0, "ERROR");
            }

            coHighList.add("ppm");
            coHighList.add("1% NIST");


        }

    }

    public void co215Extraction() throws IOException, ParseException {

        if (!"".equals(co216TextField.getText())) {

            String co216cylNum = null;

            File file = new File(bottleDirTextField.getText() + "\\" + co216TextField.getText());

            String fileText = stripPDF(bottleDirTextField.getText() + "\\" + co216TextField.getText());

            String[] fileTextArray = fileText.split("\\r\\n");

            for (int i = 0; i < fileTextArray.length; i++) {

                if (fileTextArray[i].contains("Lot Number:")) {
                    //Lot number
                    String co216Lot = fileTextArray[i].replace("Lot Number:", "");

                    co216List.add(co216Lot);

                }
                if (fileTextArray[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String co216certDate = fileTextArray[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(co216certDate);
                    co216List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Primary Master")) {

                    //Expiration Date
                    String co216expDate = fileTextArray[i + 1].replace("Expiration Date: ", "");

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(co216expDate);
                    co216List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Carbon dioxide")) {

                    //CO2 Span value
                    String co216replaceText1 = fileTextArray[i].replace("Carbon dioxide 5.5 LaserStar 15.2 %  ", "");
                    String co216span = co216replaceText1.replace(" % 1", "");

                    co216List.add(co216span);

                }

                //Cylinder number from file name
                String co216cyl = file.getName();
                //Strip out the gas type from file name
                String co216blankText1 = co216cyl.replace("CO2 16% ", "");
                //Strip out .pdf from file name
                String co216blankText2 = co216blankText1.replace(".pdf", "");
                //Actual Cylinder number from file
                co216cylNum = co216blankText2;

                if (fileTextArray[i].contains(co216cylNum)) {

                    co216List.add(co216cylNum);

                }
            }

            if (co216List.toString().contains(co216cylNum)) {

                cylNumVerificationList.add("CO2 16% Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("CO2 16% Cylinder Number Error");
            }

            if(co216List.size() < 5 ) {

                co216List.add(0, "ERROR");
            }


            co216List.add("%");
            co216List.add("1% NIST");


        }

    }

    public void egr16Extraction() throws IOException, ParseException {

        if (!"".equals(eco2TextField.getText())) {

            String eco216cylNum = null;

            File file = new File(bottleDirTextField.getText() + "\\" + eco2TextField.getText());

            String fileText = stripPDF(bottleDirTextField.getText() + "\\" + eco2TextField.getText());

            String[] fileTextArray = fileText.split("\\r\\n");

            for (int i = 0; i < fileTextArray.length; i++) {

                if (fileTextArray[i].contains("Lot Number:")) {
                    //Lot number
                    String eco216Lot = fileTextArray[i].replace("Lot Number:", "");

                    egr16List.add(eco216Lot);

                }
                if (fileTextArray[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String eco216certDate = fileTextArray[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(eco216certDate);
                    egr16List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Primary Master")) {

                    //Expiration Date
                    String eco216expDate = fileTextArray[i + 1].replace("Expiration Date: ", "");

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(eco216expDate);
                    egr16List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Carbon dioxide")) {

                    //ECO2 Span value
                    String eco216replaceText1 = fileTextArray[i].replace("Carbon dioxide 5.5 LaserStar 15.2 %  ", "");
                    String eco216span = eco216replaceText1.replace(" % 1", "");

                    egr16List.add(eco216span);

                }

                //Cylinder number from file name
                String eco216cyl = file.getName();
                //Strip out the gas type from file name
                String eco216blankText1 = eco216cyl.replace("EGR 16% ", "");
                //Strip out .pdf from file name
                String eco216blankText2 = eco216blankText1.replace(".pdf", "");
                //Actual Cylinder number from file name(Must be Fuel cyl#)
                eco216cylNum = eco216blankText2;

                if (fileTextArray[i].contains(eco216cylNum)) {

                    egr16List.add(eco216cylNum);
                }
            }

            if (egr16List.toString().contains(eco216cylNum)) {

                cylNumVerificationList.add("EGR 16% Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("EGR 16% Cylinder Number Error");
            }

            if(egr16List.size() < 5 ) {

                egr16List.add(0, "ERROR");
            }

            egr16List.add("%");
            egr16List.add("1% NIST");


        }

    }

    public void o225Extraction() throws IOException, ParseException {

        if (!"".equals(o225TextField.getText())) {

            String o225cylNum = null;


            File file = new File(bottleDirTextField.getText() + "\\" + o225TextField.getText());

            String fileText = stripPDF(bottleDirTextField.getText() + "\\" + o225TextField.getText());

            String[] fileTextArray = fileText.split("\\r\\n");

            for (int i = 0; i < fileTextArray.length; i++) {

                if (fileTextArray[i].contains("Lot Number:")) {
                    //Lot number
                    String o225Lot = fileTextArray[i].replace("Lot Number:", "");

                    o225List.add(o225Lot);

                }
                if (fileTextArray[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String o225certDate = fileTextArray[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(o225certDate);
                    o225List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Primary Master")) {

                    //Expiration Date
                    String o225expDate = fileTextArray[i + 1].replace("Expiration Date: ", "");

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(o225expDate);
                    o225List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }
                if (fileTextArray[i].contains("Oxygen")) {

                    //O2 25% Span value
                    String o225replaceText1 = fileTextArray[i].replace("Oxygen 4.3 Ultra High Purity 21.0 %  ", "");
                    String o225span = o225replaceText1.replace(" % 1", "");

                    o225List.add(o225span);

                }


                //Cylinder number from file name
                String o225cyl = file.getName();
                //Strip out the gas type from file name
                String o225blankText1 = o225cyl.replace("O2 25% ", "");
                //Strip out .pdf from file name
                String o225blankText2 = o225blankText1.replace(".pdf", "");
                //Actual Cylinder number from file name(Must be Fuel cyl#)
                o225cylNum = o225blankText2;

                if (fileTextArray[i].contains(o225cylNum)) {

                    o225List.add(o225cylNum);

                }
            }

            if (o225List.toString().contains(o225cylNum)) {

                cylNumVerificationList.add("O2 25% Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("O2 25% Cylinder Number Error");
            }

            if(o225List.size() < 5 ) {

                o225List.add(0, "ERROR");
            }


            o225List.add("%");
            o225List.add("1% NIST");


        }
    }


    public void n2Extraction() throws ParseException, IOException {

        if (!"".equals(n21TextField.getText())) {

            String n21cylNum = null;

            File file = new File(bottleDirTextField.getText() + "\\" + n21TextField.getText());

            String fileText = stripPDF(bottleDirTextField.getText() + "\\" + n21TextField.getText());

            String[] fileTextArray = fileText.split("\\r\\n");

            for (int i = 0; i < fileTextArray.length; i++) {


                if (fileTextArray[i].contains("Lot Number:")) {
                    //Lot number
                    String n21Lot = fileTextArray[i].replace("Lot Number:", "");

                    n21List.add(n21Lot);

                }
                if (fileTextArray[i].contains("Analyzed Cylinder Number(s): ")) {

                    //Analysis Cylinder
                    String n21anCyl = fileTextArray[i].replace("Analyzed Cylinder Number(s): ", "");
                    n21List.add(n21anCyl);

                }
                if (fileTextArray[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String n21certDate = fileTextArray[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(n21certDate);

                    n21List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }


                //Cylinder number from file name
                String n21cyl = file.getName();
                //Strip out the gas type from file name
                String n21blankText1 = n21cyl.replace("N2 1 ", "");
                //Strip out .pdf from file name
                String n21blankText2 = n21blankText1.replace(".pdf", "");
                //Actual Cylinder number from file name(Must be N2 cyl#)
                String n21blankText3 = n21blankText2.replace("1065 ", "");
                n21cylNum = n21blankText3;

                if (fileTextArray[i].contains(n21cylNum)) {

                    n21List.add(n21cylNum);
                }
            }

            if (n21List.toString().contains(n21cylNum)) {
                cylNumVerificationList.add("N2 1 Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("N2 1 Cylinder Number Error");
            }

        }

        if (!"".equals(n22TextField.getText())) {

            String n22cylNum = null;

            File file2 = new File(bottleDirTextField.getText() + "\\" + n22TextField.getText());

            String fileText2 = stripPDF(bottleDirTextField.getText() + "\\" + n22TextField.getText());

            String[] fileTextArray2 = fileText2.split("\\r\\n");

            for (int i = 0; i < fileTextArray2.length; i++) {

                if (fileTextArray2[i].contains("Lot Number:")) {
                    //Lot number
                    String n22Lot = fileTextArray2[i].replace("Lot Number:", "");

                    n22List.add(n22Lot);

                }
                if (fileTextArray2[i].contains("Analyzed Cylinder Number(s): ")) {

                    //Analysis Cylinder
                    String n22anCyl = fileTextArray2[i].replace("Analyzed Cylinder Number(s): ", "");

                    n22List.add(n22anCyl);

                }
                if (fileTextArray2[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String n22certDate = fileTextArray2[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(n22certDate);

                    n22List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));


                }

                //Cylinder number from file name
                String n22cyl = file2.getName();
                //Strip out the gas type from file name
                String n22blankText1 = n22cyl.replace("N2 2 ", "");
                //Strip out .pdf from file name
                String n22blankText2 = n22blankText1.replace(".pdf", "");

                String n22blankText3 = n22blankText2.replace("1065 ", "");
                //Actual Cylinder number from file name(Must be N2 cyl#)
                n22cylNum = n22blankText3;

                if (fileTextArray2[i].contains(n22cylNum)) {

                    n22List.add(n22cylNum);

                }

            }

            if (n22List.toString().contains(n22cylNum)) {
                cylNumVerificationList.add("N2 2 Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("N2 2 Cylinder Number Error");
            }
        }

        if (!"".equals(n23TextField.getText())) {

            String n23cylNum = null;

            File file3 = new File(bottleDirTextField.getText() + "\\" + n23TextField.getText());

            String fileText3 = stripPDF(bottleDirTextField.getText() + "\\" + n23TextField.getText());

            String[] fileTextArray3 = fileText3.split("\\r\\n");

            for (int i = 0; i < fileTextArray3.length; i++) {


                if (fileTextArray3[i].contains("Lot Number:")) {
                    //Lot number
                    String n23Lot = fileTextArray3[i].replace("Lot Number:", "");

                    n23List.add(n23Lot);

                }
                if (fileTextArray3[i].contains("Analyzed Cylinder Number(s): ")) {

                    //Analysis Cylinder
                    String n23anCyl = fileTextArray3[i].replace("Analyzed Cylinder Number(s): ", "");

                    n23List.add(n23anCyl);

                }
                if (fileTextArray3[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String n23certDate = fileTextArray3[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(n23certDate);

                    n23List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }


                //Cylinder number from file name
                String n23cyl = file3.getName();
                //Strip out the gas type from file name
                String n23blankText1 = n23cyl.replace("N2 3 ", "");
                //Strip out .pdf from file name
                String n23blankText2 = n23blankText1.replace(".pdf", "");

                String n23blankText3 = n23blankText2.replace("1065 ", "");
                //Actual Cylinder number from file name(Must be N2 cyl#)
                n23cylNum = n23blankText3;

                if (fileTextArray3[i].contains(n23cylNum)) {

                    n23List.add(n23cylNum);
                }
            }

            if (n23List.toString().contains(n23cylNum)) {
                cylNumVerificationList.add("N2 3 Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("N2 3 Cylinder Number Error");
            }

        }


        if (!"".equals(n24TextField.getText())) {

            String n24cylNum = null;

            File file4 = new File(bottleDirTextField.getText() + "\\" + n24TextField.getText());

            String fileText4 = stripPDF(bottleDirTextField.getText() + "\\" + n24TextField.getText());

            String[] fileTextArray4 = fileText4.split("\\r\\n");

            for (int i = 0; i < fileTextArray4.length; i++) {

                if (fileTextArray4[i].contains("Lot Number:")) {
                    //Lot number
                    String n24Lot = fileTextArray4[i].replace("Lot Number:", "");

                    n24List.add(n24Lot);

                }
                if (fileTextArray4[i].contains("Analyzed Cylinder Number(s): ")) {

                    //Analysis Cylinder
                    String n24anCyl = fileTextArray4[i].replace("Analyzed Cylinder Number(s): ", "");

                    n24List.add(n24anCyl);

                }
                if (fileTextArray4[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String n24certDate = fileTextArray4[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(n24certDate);

                    n24List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));
                }

                //Cylinder number from file name
                String n24cyl = file4.getName();
                //Strip out the gas type from file name
                String n24blankText1 = n24cyl.replace("N2 4 ", "");
                //Strip out .pdf from file name
                String n24blankText2 = n24blankText1.replace(".pdf", "");

                String n24blankText3 = n24blankText2.replace("1065 ", "");
                //Actual Cylinder number from file name(Must be N2 cyl#)
                n24cylNum = n24blankText3;

                if (fileTextArray4[i].contains(n24cylNum)) {

                    n24List.add(n24cylNum);
                }


            }

            if (n24List.toString().contains(n24cylNum)) {
                cylNumVerificationList.add("N2 4 Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("N2 4 Cylinder Number Error");
            }
        }
    }

    public void airExtraction() throws IOException, ParseException {

        if (!"".equals(air1TextField.getText())) {

            String air1cylNum = null;

            File file = new File(bottleDirTextField.getText() + "\\" + air1TextField.getText());

            String fileText = stripPDF(bottleDirTextField.getText() + "\\" + air1TextField.getText());

            String[] fileTextArray = fileText.split("\\r\\n");

            for (int i = 0; i < fileTextArray.length; i++) {

                if (fileTextArray[i].contains("Lot Number:")) {
                    //Lot number
                    String air1Lot = fileTextArray[i].replace("Lot Number:", "");
                    air1List.add(air1Lot);

                }
                if (fileTextArray[i].contains("Analyzed Cylinder Number(s): ")) {

                    //Analysis Cylinder
                    String air1anCyl = fileTextArray[i].replace("Analyzed Cylinder Number(s): ", "");
                    air1List.add(air1anCyl);


                }
                if (fileTextArray[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String air1certDate = fileTextArray[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(air1certDate);

                    air1List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }

                //Cylinder number from file name
                String air1cyl = file.getName();
                //Strip out the gas type from file name
                String air1blankText1 = air1cyl.replace("Air 1 ", "");
                //Strip out .pdf from file name
                String air1blankText2 = air1blankText1.replace(".pdf", "");

                String air1blankText3 = air1blankText2.replace("1065 ", "");
                //Actual Cylinder number from file name(Must be Air cyl#)
                air1cylNum = air1blankText3;

                if (fileTextArray[i].contains(air1cylNum)) {

                    air1List.add(air1cylNum);

                }
            }

            if (air1List.toString().contains(air1cylNum)) {
                cylNumVerificationList.add("Air 1 Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("Air 1 Cylinder Number Error");
            }

        }
        if (!"".equals(air2TextField.getText())) {

            String air2cylNum = null;

            File file2 = new File(bottleDirTextField.getText() + "\\" + air2TextField.getText());

            String fileText2 = stripPDF(bottleDirTextField.getText() + "\\" + air2TextField.getText());

            String[] fileTextArray2 = fileText2.split("\\r\\n");

            for (int i = 0; i < fileTextArray2.length; i++) {

                if (fileTextArray2[i].contains("Lot Number:")) {
                    //Lot number
                    String air2Lot = fileTextArray2[i].replace("Lot Number:", "");
                    air2List.add(air2Lot);


                }
                if (fileTextArray2[i].contains("Analyzed Cylinder Number(s): ")) {

                    //Analysis Cylinder
                    String air2anCyl = fileTextArray2[i].replace("Analyzed Cylinder Number(s): ", "");
                    air2List.add(air2anCyl);


                }
                if (fileTextArray2[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String air2certDate = fileTextArray2[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(air2certDate);

                    air2List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));


                }


                //Cylinder number from file name
                String air2cyl = file2.getName();
                //Strip out the gas type from file name
                String air2blankText1 = air2cyl.replace("Air 2 ", "");
                //Strip out .pdf from file name
                String air2blankText2 = air2blankText1.replace(".pdf", "");

                String air2blankText3 = air2blankText2.replace("1065 ", "");
                //Actual Cylinder number from file name(Must be Air cyl#)
                air2cylNum = air2blankText3;

                if (fileTextArray2[i].contains(air2cylNum)) {

                    air2List.add(air2cylNum);
                }
            }

            if (air2List.toString().contains(air2cylNum)) {
                cylNumVerificationList.add("Air 2 Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("Air 2 Cylinder Number Error");
            }


        }

        if (!"".equals(air3TextField.getText())) {

            String air3cylNum = null;

            File file3 = new File(bottleDirTextField.getText() + "\\" + air3TextField.getText());

            String fileText3 = stripPDF(bottleDirTextField.getText() + "\\" + air3TextField.getText());

            String[] fileTextArray3 = fileText3.split("\\r\\n");

            for (int i = 0; i < fileTextArray3.length; i++) {


                if (fileTextArray3[i].contains("Lot Number:")) {
                    //Lot number
                    String air3Lot = fileTextArray3[i].replace("Lot Number:", "");
                    air3List.add(air3Lot);


                }
                if (fileTextArray3[i].contains("Analyzed Cylinder Number(s): ")) {

                    //Analysis Cylinder
                    String air3anCyl = fileTextArray3[i].replace("Analyzed Cylinder Number(s): ", "");
                    air3List.add(air3anCyl);

                }
                if (fileTextArray3[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String air3certDate = fileTextArray3[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(air3certDate);

                    air3List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));

                }


                //Cylinder number from file name
                String air3cyl = file3.getName();
                //Strip out the gas type from file name
                String air3blankText1 = air3cyl.replace("Air 3 ", "");
                //Strip out .pdf from file name
                String air3blankText2 = air3blankText1.replace(".pdf", "");

                String air3blankText3 = air3blankText2.replace("1065 ", "");
                //Actual Cylinder number from file name(Must be Air cyl#)
                air3cylNum = air3blankText3;

                if (fileTextArray3[i].contains(air3cylNum)) {

                    air3List.add(air3cylNum);
                }

            }

            if (air3List.toString().contains(air3cylNum)) {
                cylNumVerificationList.add("Air 3 Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("Air 3 Cylinder Number Error");
            }
        }


        if (!"".equals(air4TextField.getText())) {

            String air4cylNum = null;

            File file4 = new File(bottleDirTextField.getText() + "\\" + air4TextField.getText());

            String fileText4 = stripPDF(bottleDirTextField.getText() + "\\" + air4TextField.getText());

            String[] fileTextArray4 = fileText4.split("\\r\\n");

            for (int i = 0; i < fileTextArray4.length; i++) {

                if (fileTextArray4[i].contains("Lot Number:")) {
                    //Lot number
                    String air4Lot = fileTextArray4[i].replace("Lot Number:", "");
                    air4List.add(air4Lot);


                }
                if (fileTextArray4[i].contains("Analyzed Cylinder Number(s): ")) {

                    //Analysis Cylinder
                    String air4anCyl = fileTextArray4[i].replace("Analyzed Cylinder Number(s): ", "");
                    air4List.add(air4anCyl);

                }
                if (fileTextArray4[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String air4certDate = fileTextArray4[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(air4certDate);

                    air4List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));


                }


                //Cylinder number from file name
                String air4cyl = file4.getName();
                //Strip out the gas type from file name
                String air4blankText1 = air4cyl.replace("Air 4 ", "");
                //Strip out .pdf from file name
                String air4blankText2 = air4blankText1.replace(".pdf", "");

                String air4blankText3 = air4blankText2.replace("1065 ", "");
                //Actual Cylinder number from file name(Must be Air cyl#)
                air4cylNum = air4blankText3;

                if (fileTextArray4[i].contains(air4cylNum)) {


                    air4List.add(air4cylNum);
                }
            }

            if (air4List.toString().contains(air4cylNum)) {
                cylNumVerificationList.add("Air 4 Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("Air 4 Cylinder Number Error");
            }


        }


    }

    public void o2100Extraction() throws IOException, ParseException {
        if (!"".equals(o21001TextField.getText())) {

            String o21001cylNum = null;

            File file = new File(bottleDirTextField.getText() + "\\" + o21001TextField.getText());

            String fileText = stripPDF(bottleDirTextField.getText() + "\\" + o21001TextField.getText());

            String[] fileTextArray = fileText.split("\\r\\n");

            for (int i = 0; i < fileTextArray.length; i++) {

                if (fileTextArray[i].contains("Lot Number:")) {
                    //Lot number
                    String o21001Lot = fileTextArray[i].replace("Lot Number:", "");
                    o21001List.add(o21001Lot);


                }
                if (fileTextArray[i].contains("Analyzed Cylinder Number(s): ")) {

                    //Analysis Cylinder
                    String o21001anCyl = fileTextArray[i].replace("Analyzed Cylinder Number(s): ", "");
                    o21001List.add(o21001anCyl);


                }
                if (fileTextArray[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String o21001certDate = fileTextArray[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date o21001newFormattedDate = dateformatter.parse(o21001certDate);
                    o21001List.add(new SimpleDateFormat("d-MMM-yyyy").format(o21001newFormattedDate));


                }
                if (fileTextArray[i].contains("Oxygen 99.")) {


                    String o21001replaceText1 = fileTextArray[i].replace("Oxygen ", "");
                    String o21001replaceText2 = o21001replaceText1.replace(" % 4", "");
                    String o21001replaceText3 = o21001replaceText2.replace(" % 3", "");

                    o21001List.add(o21001replaceText3);


                }


                //Cylinder number from file name
                String o21001cyl = file.getName();
                //Strip out the gas type from file name
                String o21001blankText1 = o21001cyl.replace("O2 100% 1 ", "");
                //Strip out .pdf from file name
                String o21001blankText2 = o21001blankText1.replace(".pdf", "");
                //Actual Cylinder number from file name(Must be N2 cyl#)
                String o21001blankText3 = o21001blankText2.replace("1065 ", "");
                o21001cylNum = o21001blankText3;

                if (fileTextArray[i].contains(o21001cylNum)) {

                    o21001List.add(o21001cylNum);
                }
            }

            if (o21001List.toString().contains(o21001cylNum)) {
                cylNumVerificationList.add("O2 100% 1 Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("O2 100% 1 Cylinder Number Error");
            }


        }

        if (!"".equals(o21002TextField.getText())) {

            String o21002cylNum = null;

            File file2 = new File(bottleDirTextField.getText() + "\\" + o21002TextField.getText());

            String fileText2 = stripPDF(bottleDirTextField.getText() + "\\" + o21002TextField.getText());

            String[] fileTextArray2 = fileText2.split("\\r\\n");

            for (int i = 0; i < fileTextArray2.length; i++) {

                if (fileTextArray2[i].contains("Lot Number:")) {
                    //Lot number
                    String o21002Lot = fileTextArray2[i].replace("Lot Number:", "");
                    o21002List.add(o21002Lot);


                }
                if (fileTextArray2[i].contains("Analyzed Cylinder Number(s): ")) {

                    //Analysis Cylinder
                    String o21002anCyl = fileTextArray2[i].replace("Analyzed Cylinder Number(s): ", "");
                    o21002List.add(o21002anCyl);


                }
                if (fileTextArray2[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String o21002certDate = fileTextArray2[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date o21002newFormattedDate = dateformatter.parse(o21002certDate);
                    o21002List.add(new SimpleDateFormat("d-MMM-yyyy").format(o21002newFormattedDate));


                }
                if (fileTextArray2[i].contains("Oxygen 99.")) {


                    String o21002replaceText1 = fileTextArray2[i].replace("Oxygen ", "");
                    String o21002replaceText2 = o21002replaceText1.replace(" % 4", "");
                    String o21002replaceText3 = o21002replaceText2.replace(" % 3", "");
                    o21002List.add(o21002replaceText3);


                }


                //Cylinder number from file name
                String o21002cyl = file2.getName();
                //Strip out the gas type from file name
                String o21002blankText1 = o21002cyl.replace("O2 100% 2 ", "");
                //Strip out .pdf from file name
                String o21002blankText2 = o21002blankText1.replace(".pdf", "");
                //Actual Cylinder number from file name(Must be N2 cyl#)
                String o21002blankText3 = o21002blankText2.replace("1065 ", "");
                o21002cylNum = o21002blankText3;

                if (fileTextArray2[i].contains(o21002cylNum)) {

                    o21002List.add(o21002cylNum);
                }
            }

            if (o21002List.toString().contains(o21002cylNum)) {
                cylNumVerificationList.add("O2 100% 2 Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("O2 100% 2 Cylinder Number Error");
            }


        }
        if (!"".equals(o21003TextField.getText())) {

            String o21003cylNum = null;

            File file3 = new File(bottleDirTextField.getText() + "\\" + o21003TextField.getText());

            String fileText3 = stripPDF(bottleDirTextField.getText() + "\\" + o21003TextField.getText());

            String[] fileTextArray3 = fileText3.split("\\r\\n");

            for (int i = 0; i < fileTextArray3.length; i++) {

                if (fileTextArray3[i].contains("Lot Number:")) {
                    //Lot number
                    String o21003Lot = fileTextArray3[i].replace("Lot Number:", "");
                    o21003List.add(o21003Lot);


                }
                if (fileTextArray3[i].contains("Analyzed Cylinder Number(s): ")) {

                    //Analysis Cylinder
                    String o21003anCyl = fileTextArray3[i].replace("Analyzed Cylinder Number(s): ", "");
                    o21003List.add(o21003anCyl);


                }
                if (fileTextArray3[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String o21003certDate = fileTextArray3[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date o21003newFormattedDate = dateformatter.parse(o21003certDate);
                    o21003List.add(new SimpleDateFormat("d-MMM-yyyy").format(o21003newFormattedDate));


                }
                if (fileTextArray3[i].contains("Oxygen 99.")) {


                    String o21003replaceText1 = fileTextArray3[i].replace("Oxygen ", "");
                    String o21003replaceText2 = o21003replaceText1.replace(" % 4", "");
                    String o21003replaceText3 = o21003replaceText2.replace(" % 3", "");
                    o21003List.add(o21003replaceText3);


                }


                //Cylinder number from file name
                String o21003cyl = file3.getName();
                //Strip out the gas type from file name
                String o21003blankText1 = o21003cyl.replace("O2 100% 3 ", "");
                //Strip out .pdf from file name
                String o21003blankText2 = o21003blankText1.replace(".pdf", "");
                //Actual Cylinder number from file name(Must be N2 cyl#)
                String o21003blankText3 = o21003blankText2.replace("1065 ", "");
                o21003cylNum = o21003blankText3;

                if (fileTextArray3[i].contains(o21003cylNum)) {

                    o21003List.add(o21003cylNum);

                }
            }

            if (o21003List.toString().contains(o21003cylNum)) {
                cylNumVerificationList.add("O2 100% 3 Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("O2 100% 3 Cylinder Number Error");
            }


        }
        if (!"".equals(o21004TextField.getText())) {

            String o21004cylNum = null;

            File file4 = new File(bottleDirTextField.getText() + "\\" + o21004TextField.getText());

            String fileText4 = stripPDF(bottleDirTextField.getText() + "\\" + o21004TextField.getText());

            String[] fileTextArray4 = fileText4.split("\\r\\n");

            for (int i = 0; i < fileTextArray4.length; i++) {

                if (fileTextArray4[i].contains("Lot Number:")) {
                    //Lot number
                    String o21004Lot = fileTextArray4[i].replace("Lot Number:", "");
                    o21004List.add(o21004Lot);


                }
                if (fileTextArray4[i].contains("Analyzed Cylinder Number(s): ")) {

                    //Analysis Cylinder
                    String o21004anCyl = fileTextArray4[i].replace("Analyzed Cylinder Number(s): ", "");
                    o21004List.add(o21004anCyl);


                }
                if (fileTextArray4[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String o21004certDate = fileTextArray4[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date o21004newFormattedDate = dateformatter.parse(o21004certDate);
                    o21004List.add(new SimpleDateFormat("d-MMM-yyyy").format(o21004newFormattedDate));


                }
                if (fileTextArray4[i].contains("Oxygen 99.")) {


                    String o21004replaceText1 = fileTextArray4[i].replace("Oxygen ", "");
                    String o21004replaceText2 = o21004replaceText1.replace(" % 4", "");
                    String o21004replaceText3 = o21004replaceText2.replace(" % 3", "");
                    o21004List.add(o21004replaceText3);


                }


                //Cylinder number from file name
                String o21004cyl = file4.getName();
                //Strip out the gas type from file name
                String o21004blankText1 = o21004cyl.replace("O2 100% 4 ", "");
                //Strip out .pdf from file name
                String o21004blankText2 = o21004blankText1.replace(".pdf", "");
                //Actual Cylinder number from file name(Must be N2 cyl#)
                String o21004blankText3 = o21004blankText2.replace("1065 ", "");
                o21004cylNum = o21004blankText3;

                if (fileTextArray4[i].contains(o21004cylNum)) {

                    o21004List.add(o21004cylNum);

                }
            }

            if (o21004List.toString().contains(o21004cylNum)) {
                cylNumVerificationList.add("O2 100% 4 Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("O2 100% 4 Cylinder Number Error");
            }

        }
    }

    public void fuelExtraction() throws ParseException, IOException {

        if (!"".equals(fuel1TextField.getText())) {

            String fuel1cylNum = null;

            File file = new File(bottleDirTextField.getText() + "\\" + fuel1TextField.getText());

            String fileText = stripPDF(bottleDirTextField.getText() + "\\" + fuel1TextField.getText());

            String[] fileTextArray = fileText.split("\\r\\n");

            for (int i = 0; i < fileTextArray.length; i++) {


                if (fileTextArray[i].contains("Lot Number:")) {
                    //Lot number
                    String fuel1Lot = fileTextArray[i].replace("Lot Number:", "");
                    fuel1List.add(fuel1Lot);


                }
                if (fileTextArray[i].contains("Analyzed Cylinder Number(s): ")) {

                    //Analysis Cylinder
                    String fuel1anCylReplaceText1 = fileTextArray[i].replace("Analyzed Cylinder Number(s): ", "");
                    String fuel1anCyl = fuel1anCylReplaceText1.replace(",", " / ");

                    fuel1List.add(fuel1anCyl);


                }
                if (fileTextArray[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String fuel1certDate = fileTextArray[i - 1];

                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(fuel1certDate);

                    fuel1List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));


                }

                //Cylinder number from file name
                String fuel1cyl = file.getName();
                //Strip out the gas type from file name
                String fuel1blankText1 = fuel1cyl.replace("Fuel 1 ", "");
                //Strip out .pdf from file name
                String fuel1blankText2 = fuel1blankText1.replace(".pdf", "");

                String fuel1blankText3 = fuel1blankText2.replace("1065 ", "");
                //Actual Cylinder number from file name(Must be Fuel cyl#)
                fuel1cylNum = fuel1blankText3;

                if (fileTextArray[i].contains(fuel1cylNum)) {

                    fuel1List.add(fuel1cylNum);

                }
            }

            if (fuel1List.toString().contains(fuel1cylNum)) {
                cylNumVerificationList.add("Fuel 1 Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("Fuel 1 Cylinder Number Error");
            }


        }

        if (!"".equals(fuel2TextField.getText())) {

            String fuel2cylNum = null;

            File file2 = new File(bottleDirTextField.getText() + "\\" + fuel2TextField.getText());

            String fileText2 = stripPDF(bottleDirTextField.getText() + "\\" + fuel2TextField.getText());

            String[] fileTextArray2 = fileText2.split("\\r\\n");

            for (int i = 0; i < fileTextArray2.length; i++) {

                if (fileTextArray2[i].contains("Lot Number:")) {
                    //Lot number
                    String fuel2Lot = fileTextArray2[i].replace("Lot Number:", "");
                    fuel2List.add(fuel2Lot);


                }
                if (fileTextArray2[i].contains("Analyzed Cylinder Number(s): ")) {

                    //Analysis Cylinder
                    String fuel2anCylReplaceText1 = fileTextArray2[i].replace("Analyzed Cylinder Number(s): ", "");
                    String fuel2anCyl = fuel2anCylReplaceText1.replace(",", " / ");
                    fuel2List.add(fuel2anCyl);


                }
                if (fileTextArray2[i].contains("Certificate Issuance Date:")) {

                    //Cert Date
                    String fuel2certDate = fileTextArray2[i - 1];
                    SimpleDateFormat dateformatter = new SimpleDateFormat("M/d/yyyy");
                    Date newFormattedDate = dateformatter.parse(fuel2certDate);
                    fuel2List.add(new SimpleDateFormat("d-MMM-yyyy").format(newFormattedDate));


                }

                //Cylinder number from file name
                String fuel2cyl = file2.getName();
                //Strip out the gas type from file name
                String fuel2blankText1 = fuel2cyl.replace("Fuel 2 ", "");
                //Strip out .pdf from file name
                String fuel2blankText2 = fuel2blankText1.replace(".pdf", "");

                String fuel2blankText3 = fuel2blankText2.replace("1065 ", "");
                //Actual Cylinder number from file name(Must be Fuel cyl#)
                fuel2cylNum = fuel2blankText3;

                if (fileTextArray2[i].contains(fuel2cylNum)) {


                    fuel2List.add(fuel2cylNum);

                }
            }

            if (fuel2List.toString().contains(fuel2cylNum)) {
                cylNumVerificationList.add("Fuel 2 Cylinder Number Verified");
            } else {
                cylNumVerificationList.add("Fuel 2 Cylinder Number Error");

            }


        }


    }


    public void clearLists() {
        nox500List.clear();
        nox2500List.clear();
        nox10kList.clear();
        no500List.clear();
        no2500List.clear();
        no10kList.clear();
        thc500List.clear();
        thc2500List.clear();
        thc10kList.clear();
        ch4500List.clear();
        ch42500List.clear();
        ch410kList.clear();
        co5000List.clear();
        co216List.clear();
        egr16List.clear();
        o225List.clear();
        cylNumVerificationList.clear();
        n21List.clear();
        n22List.clear();
        n23List.clear();
        n24List.clear();
        air1List.clear();
        air2List.clear();
        air3List.clear();
        air4List.clear();
        o21001List.clear();
        o21002List.clear();
        o21003List.clear();
        o21004List.clear();
        fuel1List.clear();
        fuel2List.clear();
    }


    public void extractFiles() throws IOException, ParseException {
try {
    nox500Extraction();
    nox2500extraction();
    nox10kExtraction();
    no500Extraction();
    no2500Extraction();
    no10kExtraction();
    thc500Extraction();
    thc2500Extraction();
    thc10kExtraction();
    ch4500Extraction();
    ch42500Extraction();
    ch410kExtraction();
    co5000Extraction();
    coHighExtraction();
    co215Extraction();
    egr16Extraction();
    o225Extraction();
    n2Extraction();
    airExtraction();
    o2100Extraction();
    fuelExtraction();

    fileDocument.close();

    FileWriter cylWriter = new FileWriter(bottleDirTextField.getText().replace("Bottles", "") + "\\" + "Cylinder Verification List.txt", false);

    for (int i = 0; i < cylNumVerificationList.size(); i++) {
        cylWriter.write(cylNumVerificationList.get(i) + "\r\n");
    }
    cylWriter.close();

    openSpreadSheet(new File(bottleDirTextField.getText().replace("Bottles", "") + "\\" + "Cylinder Verification List.txt"));
} catch (Exception e) {
    JOptionPane.showMessageDialog(null, e.toString());
}

    }

    public void writeTo1016() throws IOException, ParseException {

        String path = bottleDirTextField.getText().replace("Bottles", "");

        extractFiles();
        orderLists();

        //**************************************************************************************************************
        //Span Cylinder Loops
        //**************************************************************************************************************

        //Opening cal sheet
        File calSheetFile = new File(path + calSheetTextField.getText());

        FileInputStream calSheetInputStream = new FileInputStream(calSheetFile);
        Workbook calSheetWorkbook = WorkbookFactory.create(calSheetInputStream);
        Sheet calSheetSheet = calSheetWorkbook.getSheetAt(1);

        if (!"".equals(nox500TextField.getText())) {

            for(int i=0; i<nox500List.size(); i++) {
                excelWriter(calSheetWorkbook,calSheetSheet,1,4,i+1,nox500List.get(i));
            }
        }

        if (!"".equals(nox2500TextField.getText())) {

            for(int i=0; i<nox2500List.size(); i++) {
                excelWriter(calSheetWorkbook,calSheetSheet,1,5,i+1,nox2500List.get(i));
            }
        }

        if (!"".equals(nox10kTextField.getText())) {

            for(int i=0; i<nox10kList.size(); i++) {
                excelWriter(calSheetWorkbook,calSheetSheet,1,6,i+1,nox10kList.get(i));
            }
        }

        if (!"".equals(no500TextField.getText())) {

            for(int i=0; i<no500List.size(); i++) {
                excelWriter(calSheetWorkbook,calSheetSheet,1,7,i+1,no500List.get(i));
            }
        }

        if (!"".equals(no2500TextField.getText())) {

            for(int i=0; i<no2500List.size(); i++) {
                excelWriter(calSheetWorkbook,calSheetSheet,1,8,i+1,no2500List.get(i));
            }
        }

        if (!"".equals(no10kTextField.getText())) {

            for(int i=0; i<no10kList.size(); i++) {
                excelWriter(calSheetWorkbook,calSheetSheet,1,9,i+1,no10kList.get(i));
            }
        }

        if (!"".equals(thc500TextField.getText())) {

            for(int i=0; i<thc500List.size(); i++) {
                excelWriter(calSheetWorkbook,calSheetSheet,1,10,i+1,thc500List.get(i));
            }
        }

        if (!"".equals(thc2500TextField.getText())) {

            for(int i=0; i<thc2500List.size(); i++) {
                excelWriter(calSheetWorkbook,calSheetSheet,1,11,i+1,thc2500List.get(i));
            }
        }

        if (!"".equals(thc10kTextField.getText())) {

            for(int i=0; i<thc10kList.size(); i++) {
                excelWriter(calSheetWorkbook,calSheetSheet,1,12,i+1,thc10kList.get(i));
            }
        }

        if (!"".equals(ch4500TextField.getText())) {

            for(int i=0; i<ch4500List.size(); i++) {
                excelWriter(calSheetWorkbook,calSheetSheet,1,13,i+1,ch4500List.get(i));
            }
        }

        if (!"".equals(ch42500TextField.getText())) {

            for(int i=0; i<ch42500List.size(); i++) {
                excelWriter(calSheetWorkbook,calSheetSheet,1,14,i+1,ch42500List.get(i));
            }
        }

        if (!"".equals(ch410kTextField.getText())) {

            for(int i=0; i<ch410kList.size(); i++) {
                excelWriter(calSheetWorkbook,calSheetSheet,1,15,i+1,ch410kList.get(i));
            }
        }

        if (!"".equals(co5000TextField.getText())) {

            for(int i=0; i<co5000List.size(); i++) {
                excelWriter(calSheetWorkbook,calSheetSheet,1,16,i+1,co5000List.get(i));
            }
        }

        if (!"".equals(cohighTextField.getText())) {

            for(int i=0; i<coHighList.size(); i++) {
                excelWriter(calSheetWorkbook,calSheetSheet,1,17,i+1,coHighList.get(i));
            }
        }

        if (!"".equals(co216TextField.getText())) {

            for(int i=0; i<co216List.size(); i++) {
                excelWriter(calSheetWorkbook,calSheetSheet,1,18,i+1,co216List.get(i));
            }
        }

        if (!"".equals(eco2TextField.getText())) {

            for(int i=0; i<egr16List.size(); i++) {
                excelWriter(calSheetWorkbook,calSheetSheet,1,19,i+1,egr16List.get(i));
            }
        }

        if (!"".equals(o225TextField.getText())) {

            for(int i=0; i<o225List.size(); i++) {
                excelWriter(calSheetWorkbook,calSheetSheet,1,20,i+1,o225List.get(i));
            }
        }

        //**************************************************************************************************************
        //N2 Bottle#1
        //**************************************************************************************************************
        if (!"".equals(n21TextField.getText())) {

            //1016 N2 #1 Lot
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 24, 3, n21List.get(3));

            //1016 N2 #1 Analysis Cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 22, 3, n21List.get(1));

            //1016 N2 #1 Certification Date
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 25, 3, n21List.get(2));

            //1016 N2 #1 Cylinder Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 23, 3, n21List.get(0));
        } else {
            //1016 N2 #1 Lot
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 24, 3, "N/A");

            //1016 N2 #1 Analysis Cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 22, 3, "N/A");
        }

        //**************************************************************************************************************
        //N2 Bottle #2
        //**************************************************************************************************************
        if (!"".equals(n22TextField.getText())) {

            //1016 N2 #2 Lot Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 24, 4, n22List.get(3));

            //1016 N2 #2 Analysis Cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 22, 4, n22List.get(1));

            //1016 N2 #2 Certification Date
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 25, 4, n22List.get(2));

            //1016 N2 #2 Cylinder Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 23, 4, n22List.get(0));
        }else {
            //1016 N2 #2 Lot Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 24, 4, "N/A");

            //1016 N2 #2 Analysis Cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 22, 4, "N/A");
        }

        //**************************************************************************************************************
        //N2 Bottle #3
        //**************************************************************************************************************
        if (!"".equals(n23TextField.getText())) {

            //1016 N2 #3 Lot Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 24, 5, n23List.get(3));

            //1016 N2 #3 Analysis Cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 22, 5, n23List.get(1));

            //1016 N2 #3 Certification Date
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 25, 5, n23List.get(2));

            //1016 N2 #3 Cylinder Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 23, 5, n23List.get(0));
        }else {
            //1016 N2 #3 Lot Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 24, 5, "N/A");

            //1016 N2 #3 Analysis Cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 22, 5, "N/A");
        }

        //**************************************************************************************************************
        //N2 Bottle #4
        //**************************************************************************************************************
        if (!"".equals(n24TextField.getText())) {

            //1016 N2 #4 Lot Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 24, 6, n24List.get(3));

            //1016 N2 #3 Analysis Cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 22, 6, n24List.get(1));

            //1016 N2 #4 Certification Date
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 25, 6, n24List.get(2));

            //1016 N2 #4 Cylinder Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 23, 6, n24List.get(0));
        }else {
            //1016 N2 #4 Lot Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 24, 6, "N/A");

            //1016 N2 #3 Analysis Cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 22, 6, "N/A");
        }

        //**************************************************************************************************************
        //Air Bottle #1
        //**************************************************************************************************************
        if (!"".equals(air1TextField.getText())) {

            //1016 Air #1 Lot Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 30, 3, air1List.get(3));

            //1016 Air #1 Analysis Cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 28, 3, air1List.get(1));

            //1016 Air #1 Certification Date
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 31, 3, air1List.get(2));

            //1016 Air #1 Cylinder Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 29, 3, air1List.get(0));
        }else {
            //1016 Air #1 Lot Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 30, 3, "N/A");

            //1016 Air #1 Analysis Cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 28, 3, "N/A");
        }

        //**************************************************************************************************************
        //Air Bottle #2
        //**************************************************************************************************************
        if (!"".equals(air2TextField.getText())) {

            //1016 Air #2 Lot Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 30, 4, air2List.get(3));

            //1016 Air #2 Analysis Cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 28, 4, air2List.get(1));

            //1016 Air #2 Certification Date
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 31, 4, air2List.get(2));

            //1016 Air #2 Cylinder Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 29, 4, air2List.get(0));
        }else {
            //1016 Air #2 Lot Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 30, 4, "N/A");

            //1016 Air #2 Analysis Cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 28, 4, "N/A");
        }

        //**************************************************************************************************************
        //Air Bottle #3
        //**************************************************************************************************************
        if (!"".equals(air3TextField.getText())) {


            //1016 Air #3 Lot Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 30, 5, air3List.get(3));

            //1016 Air #3 Analysis Cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 28, 5, air3List.get(1));

            //1016 Air #3 Certification Date
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 31, 5, air3List.get(2));

            //1016 Air #3 Cylinder Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 29, 5, air3List.get(0));
        }else {
            //1016 Air #3 Lot Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 30, 5, "N/A");

            //1016 Air #3 Analysis Cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 28, 5, "N/A");
        }

        //**************************************************************************************************************
        //Air Bottle #4
        //**************************************************************************************************************
        if (!"".equals(air4TextField.getText())) {
            //1016 Air #4 Lot Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 30, 6, air4List.get(3));

            //1016 Air #4 Analysis Cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 28, 6, air4List.get(1));

            //1016 Air #4 Certification Date
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 31, 6, air4List.get(2));

            //1016 Air #4 Cylinder Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 29, 6, air4List.get(0));
        }else {
            //1016 Air #4 Lot Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 30, 6, "N/A");

            //1016 Air #4 Analysis Cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 28, 6, "N/A");
        }

        //**************************************************************************************************************
        //Fuel Bottle #1
        //**************************************************************************************************************
        if (!"".equals(fuel1TextField.getText())) {


            //1016 Fuel 1 Lot
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 36, 3, fuel1List.get(3));

            //1016 Fuel 1 analysis cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 34, 3, fuel1List.get(1));

            //1016 Fuel 1 Cert Date
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 37, 3, fuel1List.get(2));

            //1016 fuel 1 cylinder number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 35, 3, fuel1List.get(0));
        }else {
            //1016 Fuel 1 Lot
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 36, 3, "N/A");

            //1016 Fuel 1 analysis cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 34, 3, "N/A");
        }

        //**************************************************************************************************************
        //Fuel Bottle #2
        //**************************************************************************************************************
        if (!"".equals(fuel2TextField.getText())) {
            //1016 Fuel 2 Lot
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 36, 4, fuel2List.get(3));

            //1016 Fuel 2 analysis cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 34, 4, fuel2List.get(1));

            //1016 Fuel 2 Cert Date
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 37, 4, fuel2List.get(2));

            //1016 fuel 2 cylinder number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 35, 4, fuel2List.get(0));
        }else {
            //1016 Fuel 2 Lot
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 36, 4, "N/A");

            //1016 Fuel 2 analysis cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 34, 4, "N/A");
        }

        //**************************************************************************************************************
        //O2 100% Bottle #1
        //**************************************************************************************************************
        if (!"".equals(o21001TextField.getText())) {

                //1040 O2 100% #1 Lot Number
                excelWriter(calSheetWorkbook, calSheetSheet, 1, 42, 3, o21001List.get(4));

                //1040 O2 100% #1 Analysis Cylinder
                excelWriter(calSheetWorkbook, calSheetSheet, 1, 40, 3, o21001List.get(2));

                //1040 O2 100% #1 Certification Date
                excelWriter(calSheetWorkbook, calSheetSheet, 1, 43, 3, o21001List.get(3));

                //1040 O2 100% #1 Cylinder Number
                excelWriter(calSheetWorkbook, calSheetSheet, 1, 41, 3, o21001List.get(1));

                //1040 O2 100% #1 Concentration
                excelWriter(calSheetWorkbook, calSheetSheet, 1, 45, 3, o21001List.get(0));


        } else {
            //1040 O2 100% #1 Lot Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 42, 3, "N/A");
            //1040 O2 100% #1 Analysis Cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 40, 3, "N/A");
        }
        //**************************************************************************************************************
        //O2 100% Bottle #2
        //**************************************************************************************************************
        if (!"".equals(o21002TextField.getText())) {

                //1040 O2 100% #2 Lot Number
                excelWriter(calSheetWorkbook, calSheetSheet, 1, 42, 4, o21002List.get(4));

                //1040 O2 100% #2 Analysis Cylinder
                excelWriter(calSheetWorkbook, calSheetSheet, 1, 40, 4, o21002List.get(2));

                //1040 O2 100% #2 Certification Date
                excelWriter(calSheetWorkbook, calSheetSheet, 1, 43, 4, o21002List.get(3));

                //1040 O2 100% #2 Cylinder Number
                excelWriter(calSheetWorkbook, calSheetSheet, 1, 41, 4, o21002List.get(1));

                //1040 O2 100% #2 Concentration
                excelWriter(calSheetWorkbook, calSheetSheet, 1, 45, 4, o21002List.get(0));



        } else {
            //1040 O2 100% #1 Lot Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 42, 4, "N/A");
            //1040 O2 100% #2 Analysis Cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 40, 4, "N/A");
        }
        //**************************************************************************************************************
        //O2 100% Bottle #3
        //**************************************************************************************************************
        if (!"".equals(o21003TextField.getText())) {

                //1040 O2 100% #3 Lot Number
                excelWriter(calSheetWorkbook, calSheetSheet, 1, 42, 5, o21003List.get(4));

                //1040 O2 100% #3 Analysis Cylinder
                excelWriter(calSheetWorkbook, calSheetSheet, 1, 40, 5, o21003List.get(2));

                //1040 O2 100% #3 Certification Date
                excelWriter(calSheetWorkbook, calSheetSheet, 1, 43, 5, o21003List.get(3));

                //1040 O2 100% #3 Number
                excelWriter(calSheetWorkbook, calSheetSheet, 1, 41, 5, o21003List.get(1));

                //1040 O2 100% #3 Concentration
                excelWriter(calSheetWorkbook, calSheetSheet, 1, 45, 5, o21003List.get(0));



        }else {
            //1040 O2 100% #1 Lot Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 42, 5, "N/A");
            //1040 O2 100% #3 Analysis Cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 40, 5, "N/A");
        }
        //**************************************************************************************************************
        //O2 100% Bottle # 4
        //**************************************************************************************************************
        if (!"".equals(o21004TextField.getText())) {

            //1040 O2 100% #3 Lot Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 42, 6, o21003List.get(4));

            //1040 O2 100% #3 Analysis Cylinder
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 40, 6, o21003List.get(2));

            //1040 O2 100% #3 Certification Date
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 43, 6, o21003List.get(3));

            //1040 O2 100% #3 Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 41, 6, o21003List.get(1));

            //1040 O2 100% #3 Concentration
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 45, 6, o21003List.get(0));


        }else {
            //1040 O2 100% #1 Lot Number
            excelWriter(calSheetWorkbook, calSheetSheet, 1, 42, 6, "N/A");

            excelWriter(calSheetWorkbook, calSheetSheet, 1, 40, 6, "N/A");
        }


        //Closing the 1016
        calSheetInputStream.close();
        FileOutputStream calSheetOutputStream = new FileOutputStream(path + calSheetTextField.getText());
        calSheetWorkbook.write(calSheetOutputStream);
        calSheetWorkbook.close();
        calSheetOutputStream.close();

        clearLists();

        JOptionPane.showMessageDialog(null, "All Gases exported to Cal Sheet!");


        openSpreadSheet(calSheetFile);


    }

    public void excelWriter(Workbook wb, Sheet sheet, int sheetNumber, int rowNumber, int cellNumber, String cellValue) {
        wb.getSheetAt(sheetNumber);
        Row row = sheet.getRow(rowNumber);
        Cell cell = row.getCell(cellNumber);
        cell.setCellValue(cellValue);
    }

    public void openSpreadSheet(File file) throws IOException {

        //Open the File after completed
        //Check to see if desktop is supported
        if (!Desktop.isDesktopSupported()) {

            System.out.println("Desktop Not Supported");
            return;
        }


        Desktop desktop = Desktop.getDesktop();
        if (file.exists()) desktop.open(file);

    }
    
    public void orderLists() {

        //This Method will check each cylinder list to validate that it contains all the necesary data in the lists. It will then organiz them all to match the 1016 template.

        if(nox500List.size() == 7 ) {

            Collections.swap(nox500List, 3, 1);
            Collections.swap(nox500List, 4, 3);
        }

        if(nox2500List.size() == 7 ) {
            Collections.swap(nox2500List, 3, 1);
            Collections.swap(nox2500List, 4, 3);
        }

        if(nox10kList.size() == 7 ) {
            Collections.swap(nox10kList, 3, 1);
            Collections.swap(nox10kList, 4, 3);
        }

        if(no500List.size() == 7 ) {
            Collections.swap(no500List, 1, 0);
            Collections.swap(no500List, 3, 1);
            Collections.swap(no500List, 4, 3);
        }

        if(no2500List.size() == 7 ) {
            Collections.swap(no2500List, 1, 0);
            Collections.swap(no2500List, 3, 1);
            Collections.swap(no2500List, 4, 3);
        }

        if(no10kList.size() == 7 ) {
            Collections.swap(no10kList, 1, 0);
            Collections.swap(no10kList, 3, 1);
            Collections.swap(no10kList, 4, 3);
        }

        if(thc500List.size() == 7 ) {
            Collections.swap(thc500List, 1, 0);
            Collections.swap(thc500List, 3, 1);
            Collections.swap(thc500List, 4, 3);
        }

        if(thc2500List.size() == 7 ) {
            Collections.swap(thc2500List, 1, 0);
            Collections.swap(thc2500List, 3, 1);
            Collections.swap(thc2500List, 4, 3);
        }

        if(thc10kList.size() == 7 ) {
            Collections.swap(thc10kList, 1, 0);
            Collections.swap(thc10kList, 3, 1);
            Collections.swap(thc10kList, 4, 3);
        }

        if(ch4500List.size() == 7 ) {
            Collections.swap(ch4500List, 1, 0);
            Collections.swap(ch4500List, 3, 1);
            Collections.swap(ch4500List, 4, 3);
        }

        if(ch42500List.size() == 7 ) {
            Collections.swap(ch42500List, 1, 0);
            Collections.swap(ch42500List, 3, 1);
            Collections.swap(ch42500List, 4, 3);
        }

        if(ch410kList.size() == 7 ) {
            Collections.swap(ch410kList, 1, 0);
            Collections.swap(ch410kList, 3, 1);
            Collections.swap(ch410kList, 4, 3);
        }

        if(co5000List.size() == 7 ) {
            Collections.swap(co5000List, 1, 0);
            Collections.swap(co5000List, 3, 1);
            Collections.swap(co5000List, 4, 3);
        }

        if(coHighList.size() == 7 ) {
            Collections.swap(coHighList, 1, 0);
            Collections.swap(coHighList, 3, 1);
            Collections.swap(coHighList, 4, 3);
        }

        if(co216List.size() == 7 ) {
            Collections.swap(co216List, 1, 0);
            Collections.swap(co216List, 3, 1);
            Collections.swap(co216List, 4, 3);
        }

        if(egr16List.size() == 7 ) {
            Collections.swap(egr16List, 1, 0);
            Collections.swap(egr16List, 3, 1);
            Collections.swap(egr16List, 4, 3);
        }

        if(o225List.size() == 7 ) {
            Collections.swap(o225List, 1, 0);
            Collections.swap(o225List, 3, 1);
            Collections.swap(o225List, 4, 3);
        }



    }
}
