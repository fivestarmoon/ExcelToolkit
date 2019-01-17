package fsm.excelToolkit;

import java.io.File;

import javax.swing.UIManager;
import javax.swing.UIManager.LookAndFeelInfo;

import fsm.common.Log;
import fsm.excelToolkit.hmi.Window;

public class Main
{

   public static final String LogFileName_s = "ExcelToolkit.log";

   public static void main(final String[] args)
   {
      Log.Init(new File(LogFileName_s));
      Log.info("Starting ...");


      // Schedule a job for the event-dispatching thread:
      // creating and showing this application's GUI.
      javax.swing.SwingUtilities.invokeLater(new Runnable()
      {
         public void run()
         {
            try 
            {
               for (LookAndFeelInfo info : UIManager.getInstalledLookAndFeels()) {
                  if ("Nimbus".equals(info.getName())) 
                  {
                     UIManager.setLookAndFeel(info.getClassName());
                     break;
                  }
               }
            } 
            catch (Exception e) 
            {
               // If Nimbus is not available, you can set the GUI to another look and feel.
            }

            Window window = new Window();
            window.createAndShowGUI();
            if ( args.length >= 1 )
            {
               File newFile = new File(args[0]);
               if ( newFile.isFile() )
               {
                  String files[] = new String[1];
                  files[0] = newFile.getAbsolutePath();
                  window.processNewParameterFile(newFile.getAbsolutePath());
               }
            }
         }
      });
   }

}
