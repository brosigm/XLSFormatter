import java.awt.*;
import java.io.*;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;


public class FilterData {
    JFrame frame = new JFrame("XLS generátor");
    File inputFile;

    JLabel statusLabel = new JLabel("Status: ");
    JTextField statusField = new JTextField(20);

    JTextField outputFileName;
    JTextField toleranceValue;

    public FilterData(){
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(750, 140);
        frame.setResizable(false);

        PrintStream out = null;
        try {
            out = new PrintStream(new FileOutputStream("log.txt"));
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
        System.setOut(out);

        frame.setLayout(new BorderLayout());
        frame.add(setupStatus(), BorderLayout.NORTH);
        frame.add(setupInput(), BorderLayout.WEST);
        frame.add(setupOutputNameAndTolerance(), BorderLayout.EAST);
        frame.add(startButton(), BorderLayout.SOUTH);
        frame.setVisible(true);
    }

    private JPanel setupStatus(){
        JPanel statusPanel = new JPanel();
        statusField = new JTextField("Not started");
        statusField.setForeground(Color.RED);
        statusField.setEditable(false);
        statusPanel.add(statusLabel);
        statusPanel.add(statusField);
        return statusPanel;
    }

    private JPanel setupInput(){
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        fileChooser.setMultiSelectionEnabled(false);
        fileChooser.setFileFilter(new javax.swing.filechooser.FileFilter() {
            @Override
            public boolean accept(File f) {
                return f.getName().endsWith(".xls") || f.isDirectory();
            }

            @Override
            public String getDescription() {
                return "Excel file";
            }
        });

        JLabel nameOfIntputFile = new JLabel("Input file: ");
        JTextField inputFileName = new JTextField(20);
        inputFileName.setEditable(false);

        JButton chooseFile = new JButton("Choose file");
        chooseFile.addActionListener(e -> {
            int returnVal = fileChooser.showOpenDialog(frame);
            if (returnVal == JFileChooser.APPROVE_OPTION) {
                inputFile = fileChooser.getSelectedFile();
                inputFileName.setText(String.valueOf(inputFile.getAbsoluteFile()));
            }
        });

        JPanel panel = new JPanel();
        panel.add(chooseFile, BorderLayout.EAST);
        panel.add(nameOfIntputFile, BorderLayout.WEST);
        panel.add(inputFileName, BorderLayout.CENTER);
        return panel;
    }

    private JPanel setupOutputNameAndTolerance(){
        JPanel panel = new JPanel();

        JLabel nameOfOutputFile = new JLabel("Output file name: ");
        outputFileName = new JTextField(5);
        JLabel tolerance = new JLabel("Tolerance: ");
        toleranceValue = new JTextField(4);

        panel.setLayout(new GridLayout(2, 2));
        panel.add(nameOfOutputFile);
        panel.add(outputFileName);
        panel.add(tolerance);
        panel.add(toleranceValue);
        return panel;
    }

    private JPanel startButton(){
        JPanel panel = new JPanel();

        JButton start = new JButton("Start");
        start.addActionListener(e -> {
            try {
                statusField.setText("Processing...");
                statusField.setForeground(Color.YELLOW);
                System.out.println("Started reading input...");
                filterXLS(inputFile, outputFileName.getText(), Double.parseDouble(toleranceValue.getText()));
                statusField.setForeground(Color.GREEN);
                statusField.setText("Finished");
            } catch (IOException ioException) {
                System.out.println("IOException");
                ioException.printStackTrace();
            }
        });
        panel.add(start);
        return panel;
    }

    ArrayList<OneLine> lines = new ArrayList<OneLine>();
    ArrayList<OneLine> nev_lines = new ArrayList<OneLine>();
    public void filterXLS(File inputFile, String outputName, Double tolerance) throws IOException {
        {
            FileInputStream fis = null;
            try {
                fis = new FileInputStream(inputFile);
            } catch (FileNotFoundException e) {
                throw new RuntimeException(e);
            }
            HSSFWorkbook wb = null;
            try {
                wb = new HSSFWorkbook(fis);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
            HSSFSheet sheet = wb.getSheetAt(0);
            FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
            for (Row row : sheet)     //iteration over row using for each loop
            {
                OneLine line = new OneLine();
                for (Cell cell : row)    //iteration over cell using for each loop
                {
                    switch (formulaEvaluator.evaluateInCell(cell).getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:   //field that represents numeric cell type
                            if (cell.getColumnIndex() == 3) {
                                line.x = cell.getNumericCellValue();
                            } else if (cell.getColumnIndex() == 4) {
                                line.y = cell.getNumericCellValue();
                            }
                            break;
                        case Cell.CELL_TYPE_STRING:    //field that represents string cell type
                            if (cell.getColumnIndex() == 2) {
                                line.nev = cell.getStringCellValue();
                            } else if (cell.getColumnIndex() == 5) {
                                line.value = cell.getStringCellValue();
                            }
                            break;
                    }
                }
                lines.add(line);
            }

            ArrayList<OneLine> lines2 = new ArrayList<OneLine>(lines);

            for (OneLine line : lines) {
                if (line.nev.equals("H_Helyiseg_nev")) {
                    nev_lines.add(line);
                }
            }


            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet spreadsheet = workbook.createSheet(" Student Data ");
            XSSFRow row;
            Map<String, Object[]> studentData
                    = new TreeMap<String, Object[]>();

            String nev = "";
            String sorszam = "";
            String terulet = "";
            int line_NUM = 0;

            int i = 0;
            for(OneLine line : nev_lines){
                nev = "";
                sorszam = "";
                terulet = "";

                nev = line.value;
                i = 0;
                for(OneLine line2 : lines){
                    if(line2.nev.equals("H_Helyiseg_terulet") && line.getDistanceFrom(line2) < tolerance){
                        i++;
                        if(i > 2){
                            System.out.println(line_NUM + " Több terulet is van a helyiségnél!\n" + "\n   Eredeti value: " + line.value + " " + line.x + " " + line.y + "\n" + "   Új value: " + line2.value + " " + line2.x + " " + line2.y);
                        }
                        lines2.remove(line2);
                        terulet = line2.value;
                    } else if(line2.nev.equals("H_helyiseg_sorszam") && line.getDistanceFrom(line2) < tolerance){
                        i++;
                        if(i > 2){
                            System.out.println(line_NUM + " Több sorszám is van a helyiségnél!\n" + "   Eredeti value: " + line.value + " " + line.x + " " + line.y + "\n" + "   Új value: " + line2.value + " " + line2.x + " " + line2.y);
                        }
                        lines2.remove(line2);
                        sorszam = line2.value;
                    }
                }
                lines2.remove(line);

                if(i != 2){
                    studentData.put(Integer.toString(line_NUM++), new Object[]{nev, sorszam, terulet, "ERROR " + i});
                } else {
                    studentData.put(Integer.toString(line_NUM++), new Object[]{nev, sorszam, terulet});
                }


            }

            System.out.println("\nXLS jelolesek: ERROR n : n db kozeli pontot talalt a program az adott ponthoz, ami nem kettő, ezért nem tudta megfelelően kezelni");
            System.out.println("\nFel nem hasznalt adatok:");
            for(OneLine line : lines2){
                System.out.println(line.nev + " " + line.value + " " + line.x + " " + line.y);
            }




            Set<String> keyid = studentData.keySet();
            int rowid = 0;

            for (String key : keyid) {

                row = spreadsheet.createRow(rowid++);
                Object[] objectArr = studentData.get(key);
                int cellid = 0;

                for (Object obj : objectArr) {
                    Cell cell = row.createCell(cellid++);
                    cell.setCellValue((String)obj);
                }
            }

            FileOutputStream out = new FileOutputStream(
                    outputName);

            workbook.write(out);
            out.close();


        }
    }
}