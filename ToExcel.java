package csv;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.Dimension;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;

public class ToExcel {
	private static File csvFile;
    public static void main(String[] args) throws IOException {
    	JFrame frame;
        JButton button;
        JFileChooser fileChooser;
    	JButton button_two;
    	 	
    	frame = new JFrame("CSV to Excel");
        frame.setSize(1300, 400);
        frame.setVisible(true);
        
        JPanel panel = new JPanel();
        
        button = new JButton("Select File");
        button.setPreferredSize(new Dimension(90, 30)); 
        
        button_two = new JButton("Convert");
        button_two.setPreferredSize(new Dimension(130, 30)); 
        
        JLabel variableLabel = new JLabel("CSV to Excel: ", JLabel.CENTER);
         
        panel.setLayout(new GridLayout(8,3)); 
        
        panel.add(button);
        panel.add(button_two);
        panel.add(variableLabel);
        
        frame.add(panel);
        
        frame.setLocationRelativeTo(null);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setVisible(true);
        
        fileChooser = new JFileChooser();
       
        String excelFile = "data.xlsx";
        button.addActionListener(new ActionListener() {
            
			@Override
            public void actionPerformed(ActionEvent e) {
                int returnValue = fileChooser.showOpenDialog(null);
                
                if (returnValue == JFileChooser.APPROVE_OPTION) {
                	
                    File selectedFile = fileChooser.getSelectedFile();
                    System.out.println(selectedFile.getPath());
                    csvFile = fileChooser.getSelectedFile();
                    variableLabel.setText("CSV Selected!");
                    
                }
                
            }
		
        });
        
        button_two.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                
            	try (BufferedReader br = new BufferedReader(new FileReader(getCsvFile()));
                        Workbook workbook = new XSSFWorkbook();
                        FileOutputStream fileOut = new FileOutputStream(excelFile)) {

                       Sheet sheet = workbook.createSheet("Sheet1");
                       String line;
                       int rowNum = 0;

                       while ((line = br.readLine()) != null) {
                           Row row = sheet.createRow(rowNum++);
                           String[] values = line.split(",");
                          
                           for (int i = 0; i < values.length; i++) {
                               row.createCell(i).setCellValue(values[i]);             
                           }
                       }

                       workbook.write(fileOut);
                       System.out.println("CSV converted to Excel successfully!");
                       variableLabel.setText("CSV Converted!");
                   } catch (FileNotFoundException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
            }
			
        });
               
    }  
    public static File getCsvFile() {
        return csvFile;
    }
}
