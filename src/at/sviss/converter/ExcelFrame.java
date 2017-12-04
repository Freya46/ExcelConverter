/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package at.sviss.converter;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.text.BadLocationException;
import javax.swing.text.SimpleAttributeSet;
import javax.swing.text.StyleConstants;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;

/**
 *
 * @author SVISS-NEU
 */
public class ExcelFrame extends javax.swing.JFrame {

    /**
     * Creates new form ExcelFrame
     */
    final static String newline = "\n";
    public static String VERSION = "Version 1.05 - published on 11.10.2017";

    final public static int BLACK = 0;
    final public static int GREEN = 1;
    final public static int RED = 2;
    final public static int STATUS = 3;

    public static boolean DELETE_WRONG_ENTRIES = false;
    public boolean alreadyDeleted;
    public boolean deleteRow;
    public boolean containsName1String;

    public boolean headerExists;

    File importFile;
    File exportFile;

    File exportFileForOpenPath;

    Map<String, String> countryMap;
    Map<String, String> canadaProvinceMap;

    //"Familie", "Herr", "Frau", "Fr.", ...
    String[] name1Strings;

    //public static int paneRow;
    public ExcelFrame() {
        initComponents();

        alreadyDeleted = false;
        deleteRow = false;
        headerExists = false;
        containsName1String = false;

        this.setResizable(false);
        this.setTitle("ExcelConverter");

        jLabel2.setText("Fehlerhafte Zeilen werden gelöscht: " + newline + String.valueOf(ExcelFrame.DELETE_WRONG_ENTRIES));

        //paneRow = 0;
        importFilePath.setEditable(false);
        exportFilePath.setEditable(false);
        jTextArea1.setEditable(false);
        pathButton.setEnabled(false);
        exportButton.setEnabled(false);

        versionLabel.setText(VERSION);

        countryMap = new HashMap<>();
        countryMap.put("Österreich", "AT");
        countryMap.put("Deutschland", "DE");
        countryMap.put("Italien", "IT");
        countryMap.put("Belgien", "BE");
        countryMap.put("Tschechische Republik", "CZ");
        countryMap.put("Frankreich", "FR");
        countryMap.put("Ungarn", "HU");
        countryMap.put("Niederlande", "NL");
        countryMap.put("Slowakei", "SK");
        countryMap.put("Slowenien", "SI");
        countryMap.put("Spanien", "ES");
        countryMap.put("Schweiz", "CH");

        countryMap.put("Afghanistan", "AF");
        countryMap.put("Ägypten", "EG");
        countryMap.put("Argentinien", "AR");
        countryMap.put("Bosnien und Herzegowina", "BA");
        countryMap.put("Brasilien", "BR");
        countryMap.put("Bulgarien", "BG");
        countryMap.put("China", "CN");
        countryMap.put("Dänemark", "DK");
        countryMap.put("Finnland", "FI");
        countryMap.put("Griechenland", "GR");
        countryMap.put("Irland", "IE");
        countryMap.put("Irak", "IQ");
        countryMap.put("Iran", "IR");
        countryMap.put("Japan", "JP");
        countryMap.put("Kroatien", "HR");
        countryMap.put("Liechtenstein", "LI");
        countryMap.put("Norwegen", "NO");
        countryMap.put("Pakistan", "PK");
        countryMap.put("Polen", "PL");
        countryMap.put("Portugal", "PT");
        countryMap.put("Rumänien", "RO");
        countryMap.put("Russland", "RU");
        countryMap.put("Saudi-Arabien", "SA");
        countryMap.put("Schweden", "SE");
        countryMap.put("Serbien", "RS");
        countryMap.put("Großbritannien", "GB");
        countryMap.put("Oman", "OM");
        countryMap.put("Island", "IS");
        countryMap.put("Malta", "MT");
        countryMap.put("Kanada", "CA");

        countryMap.put("Vereinigte Staaten von Amerika", "US");
        countryMap.put("Estland", "EE");
        countryMap.put("Russische Föderation", "RU");
        countryMap.put("Dominica", "DM");
        countryMap.put("Australien", "AU");
        countryMap.put("Hongkong", "HK");
        countryMap.put("Litauen", "LT");
        countryMap.put("Türkei", "TR");
        countryMap.put("Thailand", "TH");
        countryMap.put("Südafrika", "ZA");
        countryMap.put("Taiwan, China", "TW");
        countryMap.put("Sri Lanka", "LK");
        countryMap.put("Korea, Republik", "KR");
        countryMap.put("Chile", "CL");
        countryMap.put("Luxemburg", "LU");
        countryMap.put("Lettland", "LV");
        countryMap.put("Ukraine", "UA");
        countryMap.put("Mexiko", "MX");
        countryMap.put("Indonesien", "ID");
        countryMap.put("Israel", "IL");
        countryMap.put("Serbia", "RS");
        countryMap.put("Neuseeland", "NZ");
        countryMap.put("Korea, Demokratische Volksrepublik", "KP");
        countryMap.put("Weißrussland", "BY");
        countryMap.put("Zypern", "CY");
        countryMap.put("Kasachstan", "KZ");
        countryMap.put("Jersey (Großbritannien)", "JE");
        countryMap.put("Vereinigten Staaten von Amerika", "US");
        countryMap.put("Kolumbien", "CO");
        countryMap.put("Indien", "IN");
        countryMap.put("Montenegro", "ME");
        countryMap.put("Vereinigte Arabische Emirate", "AE");
        countryMap.put("Mazedonien", "MK");

        String[] array = new String[countryMap.keySet().size()];

        int i = 0;
        for (Object o : countryMap.entrySet().toArray()) {
            array[i] = o.toString().replace("=", " = ");
            i++;
        }
        Arrays.sort(array);

        countryList.setListData(array);

        canadaProvinceMap = new HashMap<>();
        canadaProvinceMap.put("Alberta", "AB");
        canadaProvinceMap.put("British Columbia", "BC");
        canadaProvinceMap.put("Manitoba", "MB");
        canadaProvinceMap.put("New Brunswick", "NB");
        canadaProvinceMap.put("Neufundland und Labrador	", "NL");
        canadaProvinceMap.put("Nova Scotia", "NS");
        canadaProvinceMap.put("Ontario", "ON");
        canadaProvinceMap.put("Prince Edward Island", "PE");
        canadaProvinceMap.put("Quebec", "QC");
        canadaProvinceMap.put("Saskatchewan", "SK");
        canadaProvinceMap.put("Nordwest-Territorien", "NT");
        canadaProvinceMap.put("Nunavut", "NU");
        canadaProvinceMap.put("Yukon", "YT");

        name1Strings = new String[10];
        name1Strings[0] = "Firma";
        name1Strings[1] = "Familie";
        name1Strings[2] = "Frau";
        name1Strings[3] = "Herr";
        name1Strings[4] = "Herrn";
        name1Strings[5] = "Fam.";
        name1Strings[6] = "Hr.";
        name1Strings[7] = "Hr";
        name1Strings[8] = "Fr";
        name1Strings[9] = "Fr.";

    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        importFileChooser = new javax.swing.JFileChooser();
        exportFileChooser = new javax.swing.JFileChooser();
        jButton1 = new javax.swing.JButton();
        jButton2 = new javax.swing.JButton();
        importFilePath = new javax.swing.JTextField();
        exportFilePath = new javax.swing.JTextField();
        exportButton = new javax.swing.JButton();
        exitButton = new javax.swing.JButton();
        jScrollPane1 = new javax.swing.JScrollPane();
        jTextArea1 = new javax.swing.JTextPane();
        versionLabel = new javax.swing.JLabel();
        jScrollPane2 = new javax.swing.JScrollPane();
        countryList = new javax.swing.JList<>();
        jLabel1 = new javax.swing.JLabel();
        pathButton = new javax.swing.JButton();
        jToggleButton1 = new javax.swing.JToggleButton();
        jLabel2 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jButton1.setText("Import File Chooser");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });

        jButton2.setText("Export File Chooser");
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton2ActionPerformed(evt);
            }
        });

        exportButton.setText("Export");
        exportButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                exportButtonActionPerformed(evt);
            }
        });

        exitButton.setText("Beenden");
        exitButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                exitButtonActionPerformed(evt);
            }
        });

        jTextArea1.setMaximumSize(new java.awt.Dimension(420, 200));
        jScrollPane1.setViewportView(jTextArea1);

        versionLabel.setText("jLabel1");

        countryList.setModel(new javax.swing.AbstractListModel<String>() {
            String[] strings = { "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" };
            public int getSize() { return strings.length; }
            public String getElementAt(int i) { return strings[i]; }
        });
        jScrollPane2.setViewportView(countryList);

        jLabel1.setText("ISO-CODE2-Tabelle");

        pathButton.setText("Open Path");
        pathButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                pathButtonActionPerformed(evt);
            }
        });

        jToggleButton1.setText("Delete Error Rows");
        jToggleButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jToggleButton1ActionPerformed(evt);
            }
        });

        jLabel2.setText("jLabel2");

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addContainerGap()
                        .addComponent(versionLabel))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(75, 75, 75)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 450, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addGroup(layout.createSequentialGroup()
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                    .addComponent(jButton1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                    .addComponent(importFilePath))
                                .addGap(179, 179, 179)
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                    .addComponent(jButton2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                    .addComponent(exportFilePath))))))
                .addGap(106, 106, 106)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(18, 18, 18)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addGap(0, 118, Short.MAX_VALUE)
                                .addComponent(pathButton)
                                .addGap(18, 18, 18)
                                .addComponent(exportButton)
                                .addGap(18, 18, 18)
                                .addComponent(exitButton))
                            .addComponent(jScrollPane2)
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(jLabel1)
                                .addGap(0, 0, Short.MAX_VALUE))
                            .addComponent(jToggleButton1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                        .addGap(24, 24, 24))
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                        .addGap(75, 75, 75)
                        .addComponent(jLabel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addGap(95, 95, 95))))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addGap(47, 47, 47)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jButton1)
                    .addComponent(jButton2)
                    .addComponent(jToggleButton1))
                .addGap(18, 18, 18)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(importFilePath, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(exportFilePath, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel2))
                .addGap(17, 17, 17)
                .addComponent(jLabel1)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 220, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jScrollPane2, javax.swing.GroupLayout.PREFERRED_SIZE, 220, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 23, Short.MAX_VALUE)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(exportButton)
                    .addComponent(exitButton)
                    .addComponent(versionLabel)
                    .addComponent(pathButton))
                .addContainerGap())
        );

        pack();
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        int retVal = importFileChooser.showOpenDialog(null);

        //System.out.println(retVal);
        if (retVal == JFileChooser.APPROVE_OPTION) {
            importFile = importFileChooser.getSelectedFile();
            if (importFile.getName().endsWith(".xlsx")) {
                importFilePath.setText(importFileChooser.getSelectedFile().getName());
                exportButton.setEnabled(true);
                this.repaint();
            } else {
                JOptionPane.showMessageDialog(this, "Please select a valid file." + newline + "Only .xlsx format is supported!");
                importFile = null;
            }
        }
    }//GEN-LAST:event_jButton1ActionPerformed

    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton2ActionPerformed
        int retVal = exportFileChooser.showSaveDialog(null);

        //System.out.println(retVal);
        if (retVal == JFileChooser.APPROVE_OPTION) {
            exportFilePath.setText(exportFileChooser.getSelectedFile().getName());
            exportFile = exportFileChooser.getSelectedFile();
            this.repaint();
        }
    }//GEN-LAST:event_jButton2ActionPerformed

    private void exportButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_exportButtonActionPerformed

        if (importFile != null) {

            Date start = new Date();

            try {
                this.writeToPane(newline + newline + "Status: Begin converting " + importFile.getName() + "!" + newline + newline, ExcelFrame.STATUS);

                FileInputStream fis = new FileInputStream(importFile);
                XSSFWorkbook importWorkbook = new XSSFWorkbook(fis);
                XSSFSheet importSpreadsheet = importWorkbook.getSheetAt(0);
                Iterator<Row> importRowIterator = importSpreadsheet.iterator();

                XSSFWorkbook exportWorkbook = new XSSFWorkbook();
                XSSFSheet exportSpreadsheet = exportWorkbook.createSheet("Data");
                XSSFRow row;
                Map<Integer, Object[]> info = new TreeMap<>();
                info.put(1, new Object[]{"BusinessPartnerNumber", "Name1", "Name2", "Name3", "Name4", "CountryID", "PostalCode", "City", "AddressLine1", "Housenumber", "AddressLine2", "Tel1", "Mobile", "Fax", "Email", "Homepage", "VATID", "PersonalTaxNumber", "Eorinumber", "DeliveryInstructions", "PickupInstructions", "ProvinceISOCode"});

                Object[] olist = new Object[22];
                int i;
                int rowcount = 0;

                while (importRowIterator.hasNext()) {
                    i = 0;
                    Row nextRow = importRowIterator.next();

                    String hnr = "";
                    String str = "";

                    Cell c = nextRow.getCell(0);

                    if (c != null && c.getCellType() == Cell.CELL_TYPE_STRING && c.getStringCellValue().equals("ADDRESS_REFERENCE_NR") || c != null && c.getCellType() == Cell.CELL_TYPE_STRING && c.getStringCellValue().equals("ReferenceNr")) {
                        nextRow = importRowIterator.next();
                        headerExists = true;
                    }
                    for (int j = 0; j < 21; j++) {
                        Cell importCell = nextRow.getCell(j);

//                    if (importCell != null) {
//                        switch (importCell.getCellType()) {
//                            case Cell.CELL_TYPE_STRING:
//                                System.out.print(importCell.getStringCellValue());
//                                //olist[i] = importCell.getStringCellValue();
//                                break;
//                            case Cell.CELL_TYPE_NUMERIC:
//                                System.out.print(importCell.getNumericCellValue());
//                                //olist[i] = (int) importCell.getNumericCellValue();
//                                break;
//                            case Cell.CELL_TYPE_BLANK:
//                                System.out.println(" ");
//                                olist[i] = " ";
//                                break;
//                        }
//                    } else {
//                        olist[i] = " ";
//                    }
                        switch (j) {
                            case 0:
                                //BusinessPartnerNumber
                                if (importCell != null && importCell.getCellType() != Cell.CELL_TYPE_BLANK) {
                                    if (importCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                                        olist[0] = (int) importCell.getNumericCellValue();
                                    }
                                    if (importCell.getCellType() == Cell.CELL_TYPE_STRING) {
                                        olist[0] = importCell.getStringCellValue();
                                    }
                                } else {
                                    olist[0] = "";
                                }
                                break;
                            case 1:
                                break;
                            case 2:
                                //Name1
                                if (importCell != null && importCell.getCellType() != Cell.CELL_TYPE_BLANK) {
                                    String name1 = importCell.getStringCellValue();
                                    if (name1.equals(" ") || name1.equals("  ") || name1.equals("   ")) {
                                        if (ExcelFrame.DELETE_WRONG_ENTRIES == true && alreadyDeleted == false) {
                                            alreadyDeleted = true;
                                            deleteRow = true;

                                            int offset = 2;
                                            if (headerExists == false) {
                                                offset = 1;
                                            }

                                            this.writeToPane("Error: Name1 in Row " + (rowcount + offset) + " was too short! Possibly only whitespaces in it! Row was deleted!" + newline, ExcelFrame.RED);
                                        }
                                        if (alreadyDeleted == false) {
                                            this.writeToPane("Error: Name1 in Row " + (rowcount + 2) + " is too short! Possibly only whitespaces in it! Check manually and delete the Row if needed!" + newline, ExcelFrame.RED);
                                        }
                                        name1 = "";
                                    }
                                    //System.out.println(name1);
                                    if (stringContainsItemFromList(name1, name1Strings)) {
                                        //System.out.println("contains title");
                                        containsName1String = true;
                                        Cell c1 = nextRow.getCell(2);
                                        Cell c2 = nextRow.getCell(3);
                                        olist[3] = "";
                                        
                                        //System.out.println(c1.getStringCellValue());

                                        if (c1 != null && c1.getCellType() == Cell.CELL_TYPE_STRING) {
                                            name1 = c1.getStringCellValue();
                                            //System.out.print(nextRow.getRowNum() + ": " + name1);
                                        } else {
                                            name1 = "";
                                        }
                                        if (c2 != null && c2.getCellType() == Cell.CELL_TYPE_STRING) {
                                            if (name1.equals("")) {
                                                name1 = c2.getStringCellValue();
                                                //System.out.println(name1);
                                            } else {
                                                name1 = name1 + " " + c2.getStringCellValue();
                                                //System.out.println(name1);
                                            }
                                        }
                                    }

                                    olist[1] = name1;
                                } else {
                                    olist[1] = "";
                                }
                                break;
                            case 3:
                                //Name2
                                if (containsName1String != true) {
                                    if (importCell != null && importCell.getCellType() != Cell.CELL_TYPE_BLANK) {
                                        //System.out.println(importCell.getRowIndex() + " " + importCell.getCellType());
                                        olist[2] = importCell.getStringCellValue();
                                    } else {
                                        olist[2] = "";
                                    }
                                }
                                break;
                            case 4:
                                //Name3
                                if (containsName1String != true) {
                                    if (importCell != null && importCell.getCellType() != Cell.CELL_TYPE_BLANK) {
                                        olist[3] = importCell.getStringCellValue();
                                    } else {
                                        olist[3] = "";
                                    }
                                } else {
                                    if (importCell != null && importCell.getCellType() != Cell.CELL_TYPE_BLANK) {
                                        olist[2] = importCell.getStringCellValue();
                                    } else {
                                        olist[2] = "";
                                    }
                                }

                                break;
                            case 5:
                                //AddressLine1
                                if (importCell != null && importCell.getCellType() != Cell.CELL_TYPE_BLANK) {
                                    //System.out.print("1 " + importCell.getStringCellValue() + " ");
                                    //System.out.println(importCell.getRowIndex());
                                    str = importCell.getStringCellValue();

                                    //If the addressline1 consists only of whitespaces " " 
                                    if (str.equals(" ") || str.equals("  ") || str.equals("   ")) {
                                        str = "";
                                        if (ExcelFrame.DELETE_WRONG_ENTRIES == true && alreadyDeleted == false) {
                                            alreadyDeleted = true;
                                            deleteRow = true;

                                            int offset = 2;
                                            if (headerExists == false) {
                                                offset = 1;
                                            }
                                            this.writeToPane("Error: No Addressline1 was entered in Row " + (rowcount + offset) + "! Row was deleted!" + newline, ExcelFrame.RED);
                                        }
                                        if (alreadyDeleted == false) {
                                            this.writeToPane("Error: No Addressline1 was entered in Row " + (rowcount + 2) + "! Please check manually and delete the Row if needed." + newline, ExcelFrame.RED);
                                        }
                                    }
                                    //System.out.println(importCell.getStringCellValue() + " " + str);
                                } else {
                                    str = "";
                                    //System.out.println("Error: No Addressline1");
                                    //jTextArea1.
                                    //jTextArea1.append("Error: No Addressline1 was entered in Row " + (rowcount + 1) + "! Please check manually and delete the Row if needed." + newline);
                                    if (ExcelFrame.DELETE_WRONG_ENTRIES == true && alreadyDeleted == false) {
                                        alreadyDeleted = true;
                                        deleteRow = true;

                                        int offset = 2;
                                        if (headerExists == false) {
                                            offset = 1;
                                        }
                                        this.writeToPane("Error: No Addressline1 was entered in Row " + (rowcount + offset) + "! Row was deleted!" + newline, ExcelFrame.RED);
                                    }
                                    if (alreadyDeleted == false) {
                                        this.writeToPane("Error: No Addressline1 was entered in Row " + (rowcount + 2) + "! Please check manually and delete the Row if needed." + newline, ExcelFrame.RED);
                                    }
                                }
                                break;
                            case 6:
                                //Housenumber part 1
                                if (importCell != null && importCell.getCellType() != Cell.CELL_TYPE_BLANK) {
                                    //System.out.print(" 2" + importCell.getStringCellValue());
                                    if (importCell.getCellType() == Cell.CELL_TYPE_STRING) {
                                        hnr = importCell.getStringCellValue();
                                    }
                                    if (importCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                                        hnr = String.valueOf((int) importCell.getNumericCellValue());
                                    }
                                } else {
                                    hnr = "";
                                }
                                break;
                            case 7:
                                //Housenumber part 2
                                if (importCell != null && importCell.getCellType() != Cell.CELL_TYPE_BLANK) {
                                    //System.out.print(" 3" + importCell.getStringCellValue());
                                    if (importCell.getCellType() == Cell.CELL_TYPE_STRING) {
                                        hnr = hnr.concat(" ");
                                        hnr = hnr.concat(importCell.getStringCellValue());
                                    }
                                    if (importCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                                        hnr = hnr.concat(" ");
                                        hnr = hnr.concat(String.valueOf((int) importCell.getNumericCellValue()));
                                    } else {
                                        hnr = hnr.concat("");
                                    }
                                }
                                break;
                            case 8:
                                //PostalCode
                                if (importCell != null && importCell.getCellType() != Cell.CELL_TYPE_BLANK) {
                                    if (importCell.getCellType() == Cell.CELL_TYPE_STRING) {
                                        olist[6] = importCell.getStringCellValue();
                                    }
                                    if (importCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                                        olist[6] = (int) importCell.getNumericCellValue();
                                    }
                                }
                                break;
                            case 9:
                                //City
                                if (importCell != null && importCell.getCellType() != Cell.CELL_TYPE_BLANK) {
                                    olist[7] = importCell.getStringCellValue();
                                } else {
                                    olist[7] = "";
                                    //jTextArea1.append("Warning: No City was entered in Row " + (rowcount + 1) + "! Please check manually and delete the Row if needed." + newline);
                                    this.writeToPane("Warning: No City was entered in Row " + (rowcount + 1) + "! Please check manually and delete the Row if needed." + newline, ExcelFrame.BLACK);
                                }
                                break;
                            case 10:
                                //ProvinceISOCode for Canada!
                                //ADDRESS_DETAILS -> meist nicht verwendet (99.99999999%)
                                Cell c1 = nextRow.getCell(12);
                                if (c1 != null && c1.getCellType() == Cell.CELL_TYPE_STRING && c1.getStringCellValue().equals("Kanada")) {
                                    if (importCell != null && importCell.getCellType() == Cell.CELL_TYPE_STRING) {
                                        String provinceCode2 = canadaProvinceMap.get(importCell.getStringCellValue());
                                        if (provinceCode2 == null || provinceCode2.equals(" ")) {
                                            provinceCode2 = "";
                                        }
                                        olist[21] = "CA-" + provinceCode2;
                                    }

                                }
                                break;
                            case 11:
                                break;
                            case 12:
                                //CountryID
                                if (importCell != null && importCell.getCellType() != Cell.CELL_TYPE_BLANK) {
                                    String isoCode2 = countryMap.get(importCell.getStringCellValue());
                                    if (isoCode2 != null) {
                                        olist[5] = isoCode2;

                                        Cell provinceCodeCell_pc;
                                        Cell provinceCodeCell_city;
                                        switch (isoCode2) {

                                            case "CA":
                                                provinceCodeCell_pc = nextRow.getCell(8);
                                                if (provinceCodeCell_pc != null && provinceCodeCell_pc.getCellType() == Cell.CELL_TYPE_STRING) {
                                                    //System.out.println(isoCode2 + " " + provinceCodeCell.getStringCellValue());
                                                    String pic2 = (provinceCodeCell_pc.getStringCellValue().split(" "))[0];
                                                    //System.out.println((rowcount + 2) + " " + pic2);

                                                    if (pic2.length() == 2) {
                                                        olist[21] = "CA-" + pic2;
                                                        this.writeToPane("Warning: Setting ProvinceISOCode in Row " + (rowcount + 1) + " to " + pic2 + " was '" + nextRow.getCell(21) + "'." + newline, ExcelFrame.BLACK);
                                                    }

                                                }
                                                provinceCodeCell_city = nextRow.getCell(9);
                                                if (provinceCodeCell_city != null && provinceCodeCell_city.getCellType() == Cell.CELL_TYPE_STRING) {
                                                    String[] split = provinceCodeCell_city.getStringCellValue().split(" ");
                                                    String pic2 = split[split.length - 1];
                                                    //System.out.println(pic2);

                                                    if (pic2.length() > 2) {
                                                        pic2 = pic2.toLowerCase();
                                                        pic2 = pic2.substring(0, 1).toUpperCase() + pic2.substring(1);
                                                        pic2 = canadaProvinceMap.get(pic2);
                                                    }
                                                    if (pic2 != null && pic2.length() == 2) {
                                                        olist[21] = "CA-" + pic2;
                                                        this.writeToPane("Warning: Setting ProvinceISOCode in Row " + (rowcount + 1) + " to " + pic2 + " was '" + nextRow.getCell(21) + "'." + newline, ExcelFrame.BLACK);
                                                    }
                                                }
                                                break;
                                            case "US":
                                                provinceCodeCell_pc = nextRow.getCell(8);
                                                if (provinceCodeCell_pc != null && provinceCodeCell_pc.getCellType() == Cell.CELL_TYPE_STRING) {
                                                    String pic2 = provinceCodeCell_pc.getStringCellValue().substring(0, 2);
                                                    //System.out.println(isoCode2 + " " + pic2);
                                                    olist[21] = "US-" + pic2;
                                                    this.writeToPane("Warning: Setting ProvinceISOCode in Row " + (rowcount + 1) + " to " + pic2 + " was '" + nextRow.getCell(21) + "'." + newline, ExcelFrame.BLACK);
                                                }
                                                break;
                                        }

                                    } else {
                                        olist[5] = importCell.getStringCellValue();
                                        //jTextArea1.append("Error: System doesn't know the ISO-CODE2 for '" + importCell.getStringCellValue() + "'! Please change it manually! Row: " + importCell.getRowIndex() + 1 + newline);
                                        if (ExcelFrame.DELETE_WRONG_ENTRIES == true && alreadyDeleted == false) {
                                            alreadyDeleted = true;
                                            deleteRow = true;

                                            int offset = 1;
                                            if (headerExists == false) {
                                                offset = 0;
                                            }
                                            this.writeToPane("Error: System doesn't know the ISO-CODE2 for '" + importCell.getStringCellValue() + "'! Row was deleted! Row: " + (importCell.getRowIndex() + offset) + newline, ExcelFrame.RED);
                                        }
                                        if (alreadyDeleted == false) {
                                            this.writeToPane("Error: System doesn't know the ISO-CODE2 for '" + importCell.getStringCellValue() + "'! Please change it manually! Row: " + importCell.getRowIndex() + 1 + newline, ExcelFrame.RED);
                                        }
                                    }
                                }
                                break;
                            case 13:
                                //Telnum
                                boolean flag = false;
                                if (importCell != null && importCell.getCellType() != Cell.CELL_TYPE_BLANK) {
                                    //System.out.print((rowcount + 2));
                                    if (importCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                                        olist[11] = (int) importCell.getNumericCellValue();
                                    } else if (importCell.getCellType() == Cell.CELL_TYPE_STRING) {
                                        olist[11] = importCell.getStringCellValue();
                                    }
                                    //System.out.println(" " + olist[11]);
                                    flag = true;
                                }
                                if (flag == false) {
                                    if (importCell != null && importCell.getCellType() == Cell.CELL_TYPE_BLANK) {
                                        //System.out.println((rowcount + 2) + " blank");
                                        olist[11] = "";
                                    } else {
                                        //System.out.println((rowcount + 2) + " null");
                                        olist[11] = "";
                                    }
                                }
                                break;
                            case 14:
                                //Fax
                                flag = false;
                                if (importCell != null && importCell.getCellType() != Cell.CELL_TYPE_BLANK) {
                                    //System.out.print((rowcount + 2));
                                    if (importCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                                        olist[13] = (int) importCell.getNumericCellValue();
                                    } else if (importCell.getCellType() == Cell.CELL_TYPE_STRING) {
                                        olist[13] = importCell.getStringCellValue();
                                    }
                                    //System.out.println(" " + olist[11]);
                                    flag = true;
                                }
                                if (flag == false) {
                                    if (importCell != null && importCell.getCellType() == Cell.CELL_TYPE_BLANK) {
                                        //System.out.println((rowcount + 2) + " blank");
                                        olist[13] = "";
                                    } else {
                                        //System.out.println((rowcount + 2) + " null");
                                        olist[13] = "";
                                    }
                                }
                                break;
                            case 15:
                                //Mobile
                                flag = false;
                                if (importCell != null && importCell.getCellType() != Cell.CELL_TYPE_BLANK) {
                                    //System.out.print((rowcount + 2));
                                    if (importCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                                        olist[12] = (int) importCell.getNumericCellValue();
                                    } else if (importCell.getCellType() == Cell.CELL_TYPE_STRING) {
                                        olist[12] = importCell.getStringCellValue();
                                    }
                                    //System.out.println(" " + olist[11]);
                                    flag = true;
                                }
                                if (flag == false) {
                                    if (importCell != null && importCell.getCellType() == Cell.CELL_TYPE_BLANK) {
                                        //System.out.println((rowcount + 2) + " blank");
                                        olist[12] = "";
                                    } else {
                                        //System.out.println((rowcount + 2) + " null");
                                        olist[12] = "";
                                    }
                                }
                                break;
                            case 16:
                                //Email
                                flag = false;
                                if (importCell != null && importCell.getCellType() != Cell.CELL_TYPE_BLANK) {
                                    //System.out.print((rowcount + 2));
                                    if (importCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                                        olist[14] = (int) importCell.getNumericCellValue();
                                    } else if (importCell.getCellType() == Cell.CELL_TYPE_STRING) {
                                        olist[14] = importCell.getStringCellValue();
                                    }
                                    //System.out.println(" " + olist[11]);
                                    flag = true;
                                }
                                if (flag == false) {
                                    if (importCell != null && importCell.getCellType() == Cell.CELL_TYPE_BLANK) {
                                        //System.out.println((rowcount + 2) + " blank");
                                        olist[14] = "";
                                    } else {
                                        //System.out.println((rowcount + 2) + " null");
                                        olist[14] = "";
                                    }
                                }
                                break;
                            case 17:
                                //Website
                                flag = false;
                                if (importCell != null && importCell.getCellType() != Cell.CELL_TYPE_BLANK) {
                                    //System.out.print((rowcount + 2));
                                    if (importCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                                        olist[15] = (int) importCell.getNumericCellValue();
                                    } else if (importCell.getCellType() == Cell.CELL_TYPE_STRING) {
                                        olist[15] = importCell.getStringCellValue();
                                    }
                                    //System.out.println(" " + olist[11]);
                                    flag = true;
                                }
                                if (flag == false) {
                                    if (importCell != null && importCell.getCellType() == Cell.CELL_TYPE_BLANK) {
                                        //System.out.println((rowcount + 2) + " blank");
                                        olist[15] = "";
                                    } else {
                                        //System.out.println((rowcount + 2) + " null");
                                        olist[15] = "";
                                    }
                                }
                                break;
                            case 18:
                                break;
                            case 19:
                                //TaxId
                                flag = false;
                                if (importCell != null && importCell.getCellType() != Cell.CELL_TYPE_BLANK) {
                                    //System.out.print((rowcount + 2));
                                    if (importCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                                        olist[16] = (int) importCell.getNumericCellValue();
                                    } else if (importCell.getCellType() == Cell.CELL_TYPE_STRING) {
                                        olist[16] = importCell.getStringCellValue();
                                    }
                                    //System.out.println(" " + olist[11]);
                                    flag = true;
                                }
                                if (flag == false) {
                                    if (importCell != null && importCell.getCellType() == Cell.CELL_TYPE_BLANK) {
                                        //System.out.println((rowcount + 2) + " blank");
                                        olist[16] = "";
                                    } else {
                                        //System.out.println((rowcount + 2) + " null");
                                        olist[16] = "";
                                    }
                                }
                                break;
                            case 20:
                                //DeliveryInstructions
                                if (importCell != null && importCell.getCellType() != Cell.CELL_TYPE_BLANK) {
                                    if (importCell.getCellType() == Cell.CELL_TYPE_STRING) {
                                        olist[19] = importCell.getStringCellValue();
                                    }
                                } else {
                                    olist[19] = "";
                                }
                                break;
                        }

                        //System.out.println(str + " " + hnr);
                        //System.out.println(str);
                        olist[8] = str;
                        //System.out.println(olist[8]);
                        olist[9] = hnr;

                        //Leere Felder
                        if (olist[4] == null) {
                            olist[4] = "";
                        }
                        if (olist[10] == null) {
                            olist[10] = "";
                        }
                        if (olist[12] == null) {
                            olist[12] = "";
                        }
                        if (olist[13] == null) {
                            olist[13] = "";
                        }
                        if (olist[14] == null) {
                            olist[14] = "";
                        }
                        if (olist[15] == null) {
                            olist[15] = "";
                        }
                        if (olist[16] == null) {
                            olist[16] = "";
                        }
                        if (olist[17] == null) {
                            olist[17] = "";
                        }
                        if (olist[18] == null) {
                            olist[18] = "";
                        }
                        if (olist[20] == null) {
                            olist[20] = "";
                        }
                        if (olist[21] == null) {
                            olist[21] = "";
                        }
                        //System.out.print(" " + str + " " + hnr);
                        if (hnr.length() >= 11) {
                            //jTextArea1.append("Warning: Check if Row " + (rowcount + 1) + " is correct. Housenumber was longer then 10 characters!" + newline);
                            this.writeToPane("Warning: Check if Row " + (rowcount + 2) + " is correct. Housenumber was longer then 10 characters!" + newline, ExcelFrame.BLACK);
                            //String sub = hnr.substring(10);
                            //System.out.print("sub: " + sub + " hnr: " + hnr);
                            //hnr = hnr.replace(sub, "");
                            //System.out.print(" hnr2: " + hnr);
                            str = str.concat(" ");
                            //System.out.print(" str: " + str);
                            str = str.concat(hnr);
                            //System.out.print(" str2: " + str);
                            //hnr = sub;
                            //System.out.println(str + " " + hnr);
                            hnr = "";
                        }

                        //System.out.print(" - ");
                        i = i + 1;
                    }
                    //System.out.println();
                    if (deleteRow == false) {
                        rowcount = rowcount + 1;
                        info.put(rowcount + 1, olist);
                    } else {

                    }
                    //info.put(rowcount + 1, olist);
                    olist = new Object[22];
                    alreadyDeleted = false;
                    deleteRow = false;
                    containsName1String = false;
                }

                Set< Integer> keyid = info.keySet();
                int rowid = 0;
                for (Integer key : keyid) {
                    row = exportSpreadsheet.createRow(rowid++);
                    Object[] objectArr = info.get(key);
                    int cellid = 0;
                    for (Object obj : objectArr) {
                        Cell cell = row.createCell(cellid++);
                        cell.setCellValue(String.valueOf(obj));
                    }
                }

                FileOutputStream fos;
                //System.out.println(exportFile.getName());
                if (exportFile == null) {
                    String date = new SimpleDateFormat("dd-MM-yyyy_HHmm").format(new Date());

                    //System.out.println(importFile.getParent() + "\\export_" + date);
                    exportFile = new File(importFile.getPath().substring(0, importFile.getPath().length() - 5) + "_converted_" + date);
                    //System.out.println(exportFile.getPath());
                    //exportFile.createNewFile();
                }

                fos = new FileOutputStream(exportFile + ".xlsx");

                exportWorkbook.write(fos);
                fos.close();
                //System.out.println(exportFile.getName() + ".xlsx written successfully: " + rowcount + " rows exported!");
                //jTextArea1.append(newline + exportFile.getName() + ".xlsx written successfully: " + rowcount + " rows exported!" + newline + newline);

                Date end = new Date();
                double elapsedTime = (((double) end.getTime() - start.getTime()) / 1000);

                this.writeToPane(newline + exportFile.getName() + ".xlsx written successfully: " + rowcount + " rows exported in " + elapsedTime + " Seconds!" + newline + newline, ExcelFrame.GREEN);

                fis.close();
                pathButton.setEnabled(true);
                exportFileForOpenPath = exportFile;
                exportFilePath.setText("");
                exportFile = null;
            } catch (FileNotFoundException ex) {
                Logger.getLogger(ExcelFrame.class.getName()).log(Level.SEVERE, null, ex);
            } catch (IOException ex) {
                Logger.getLogger(ExcelFrame.class.getName()).log(Level.SEVERE, null, ex);
            } catch (BadLocationException ex) {
                Logger.getLogger(ExcelFrame.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }//GEN-LAST:event_exportButtonActionPerformed

    private void exitButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_exitButtonActionPerformed
        System.exit(0);
    }//GEN-LAST:event_exitButtonActionPerformed

    private void pathButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_pathButtonActionPerformed
        try {
            //Desktop.getDesktop().open(new File(exportFile.getParent()));
            //Runtime.getRuntime().exec("explorer.exe /select " + exportFile.getPath());
            //System.out.println(exportFile.getPath() + ".xlsx");
            if (exportFileForOpenPath.getPath().contains(" ")) {
                Process p = new ProcessBuilder("explorer.exe", "\"/select," + exportFileForOpenPath.getPath() + ".xlsx").start();
            } else {
                Process p = new ProcessBuilder("explorer.exe", "/select," + exportFileForOpenPath.getPath() + ".xlsx").start();
            }

        } catch (IOException ex) {
            Logger.getLogger(ExcelFrame.class.getName()).log(Level.SEVERE, null, ex);
        }
        //exportFile = null;
    }//GEN-LAST:event_pathButtonActionPerformed

    private void jToggleButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jToggleButton1ActionPerformed
        ExcelFrame.DELETE_WRONG_ENTRIES = !ExcelFrame.DELETE_WRONG_ENTRIES;
        jLabel2.setText("Fehlerhafte Zeilen werden gelöscht: " + newline + String.valueOf(ExcelFrame.DELETE_WRONG_ENTRIES));
    }//GEN-LAST:event_jToggleButton1ActionPerformed

    public void writeToPane(String message, int color) throws BadLocationException {
        SimpleAttributeSet c = new SimpleAttributeSet();
        StyleConstants.setFontFamily(c, "Times New Roman");
        switch (color) {
            case 0:
                StyleConstants.setForeground(c, Color.BLACK);
                break;
            case 1:
                StyleConstants.setForeground(c, Color.BLACK);
                StyleConstants.setBackground(c, Color.GREEN);
                StyleConstants.setBold(c, true);
                break;
            case 2:
                StyleConstants.setForeground(c, Color.RED);
                StyleConstants.setBold(c, true);
                break;
            case 3:
                StyleConstants.setForeground(c, Color.WHITE);
                StyleConstants.setBackground(c, Color.BLACK);
                StyleConstants.setBold(c, true);
                break;
        }

        jTextArea1.getDocument().insertString(jTextArea1.getDocument().getLength(), message, c);
    }

    public static boolean stringContainsItemFromList(String inputStr, String[] items) {
        return Arrays.stream(items).parallel().anyMatch(inputStr::equals);
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(ExcelFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(ExcelFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(ExcelFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(ExcelFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new ExcelFrame().setVisible(true);
            }
        });
    }


    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JList<String> countryList;
    private javax.swing.JButton exitButton;
    private javax.swing.JButton exportButton;
    private javax.swing.JFileChooser exportFileChooser;
    private javax.swing.JTextField exportFilePath;
    private javax.swing.JFileChooser importFileChooser;
    private javax.swing.JTextField importFilePath;
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JTextPane jTextArea1;
    private javax.swing.JToggleButton jToggleButton1;
    private javax.swing.JButton pathButton;
    private javax.swing.JLabel versionLabel;
    // End of variables declaration//GEN-END:variables
}
