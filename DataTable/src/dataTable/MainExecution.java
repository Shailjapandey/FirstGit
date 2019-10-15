package dataTable;

import java.awt.EventQueue;

public class MainExecution
{
  public MainExecution() {}
  
  public static void main(String[] args) {
    EventQueue.invokeLater(new Runnable() {
      public void run() {
        try {
          UIWindow window = new UIWindow();
          window.frame.setVisible(true);
        } catch (Exception e) {
          e.printStackTrace();
        }
      }
    });
  }
}