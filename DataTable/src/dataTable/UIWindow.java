package dataTable;

import java.awt.Container;
import java.awt.event.ActionEvent;
import java.util.List;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JTextField;

public class UIWindow
{
  JFrame frame;
  JLabel folLabel;
  JLabel fileLabel;
  JTextField txtFolderPath;
  JTextField txtFilePath;
  JButton btnFolder;
  JButton btnFile;
  JButton startBtn;
  String folderPath;
  String filePath;
  String folderPath1;
  String filePath1;
  String getCurrentDirectory;
  List<String> header = new java.util.ArrayList();
  List<String> data = new java.util.ArrayList();
  
  BussinessLogic bLogic = new BussinessLogic();
  
  public UIWindow()
  {
    initialize();
  }
  



  private void initialize()
  {
    frame = new JFrame("DataTable Creator");
    frame.setBounds(100, 100, 900, 300);
    frame.setDefaultCloseOperation(3);
    frame.getContentPane().setLayout(null);
    


    folLabel = new JLabel("Select XML Repository ");
    folLabel.setBounds(10, 41, 300, 30);
    folLabel.setFont(new java.awt.Font("Courier New", 1, 15));
    frame.getContentPane().add(folLabel);
    


    fileLabel = new JLabel("Select DataTable Format File ");
    fileLabel.setBounds(10, 81, 300, 30);
    fileLabel.setFont(new java.awt.Font("Courier New", 1, 15));
    frame.getContentPane().add(fileLabel);
    

    txtFolderPath = new JTextField();
    txtFolderPath.setBounds(280, 41, 400, 25);
    frame.getContentPane().add(txtFolderPath);
    

    txtFilePath = new JTextField();
    txtFilePath.setBounds(280, 81, 400, 25);
    frame.getContentPane().add(txtFilePath);
    

    btnFolder = new JButton("Browse");
    btnFolder.setBounds(700, 41, 87, 23);
    frame.getContentPane().add(btnFolder);
    getFolderName();
    

    btnFile = new JButton("Browse");
    btnFile.setBounds(700, 81, 87, 23);
    frame.getContentPane().add(btnFile);
    getFileName();
    

    startBtn = new JButton("START");
    startBtn.setBounds(400, 150, 87, 23);
    frame.getContentPane().add(startBtn);
    Start();
  }
  



  public String getFolderName()
  {
    btnFolder.addActionListener(new java.awt.event.ActionListener()
    {
      public void actionPerformed(ActionEvent e) {
        JFileChooser fileChooser = new JFileChooser("C:/");
        fileChooser.setFileSelectionMode(1);
        fileChooser.setAcceptAllFileFilterUsed(false);
        int rVal = fileChooser.showOpenDialog(null);
        if (rVal == 0) {
          txtFolderPath.setText(fileChooser.getSelectedFile().toString());
          folderPath = fileChooser.getSelectedFile().toString().replace("\\", "/");
          getCurrentDirectory = fileChooser.getCurrentDirectory().toString().replace("\\", "/");
          System.out.println("Select FolderPath " + folderPath);
        }
      }
    });
    return folderPath;
  }
  


  public String getFileName()
  {
    btnFile.addActionListener(new java.awt.event.ActionListener()
    {
      public void actionPerformed(ActionEvent e) {
        JFileChooser fileChooser1 = new JFileChooser("C:/");
        fileChooser1.setFileSelectionMode(0);
        fileChooser1.setAcceptAllFileFilterUsed(false);
        int rVal = fileChooser1.showOpenDialog(null);
        if (rVal == 0) {
          txtFilePath.setText(fileChooser1.getSelectedFile().toString());
          filePath = fileChooser1.getSelectedFile().toString().replace("\\", "/");
          System.out.println("Select FilePath " + filePath);
        }
        
      }
    });
    return filePath;
  }
  

  public void Start()
  {
    System.out.println("Clicked On Start Button  ");
    startBtn.addActionListener(new java.awt.event.ActionListener()
    {

      public void actionPerformed(ActionEvent e)
      {
        try
        {
          java.util.Map<String, List<String>> excelValue = bLogic.readExcelFile(filePath, 0);
          
          header = ((List)excelValue.get("headerlist"));
          data = ((List)excelValue.get("datalist"));
          


          bLogic.createExcel(getCurrentDirectory + "/FinalDataTable.xlsx", "QuoteCreation", header);
          


          java.util.ArrayList<String> fileListName = bLogic.readFilename(folderPath);
          

          java.util.ArrayList<String> missingFile = bLogic.compareFiles(folderPath);
          

          bLogic.writeDatainExitingExcel(fileListName, data, getCurrentDirectory + "/FinalDataTable.xlsx");
          

          bLogic.writeDatainExitingExcel(missingFile, getCurrentDirectory + "/FinalDataTable.xlsx", "MissingFile");
          javax.swing.JOptionPane.showMessageDialog(frame, "Data Table is created");
        }
        catch (Exception e1) {
          e1.printStackTrace();
        }
      }
    });
  }
}