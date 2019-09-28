package fsm.excelToolkit;

import java.io.File;

import javax.swing.UIManager;
import javax.swing.UIManager.LookAndFeelInfo;

import fsm.common.Log;
import fsm.excelToolkit.hmi.Window;

public class Main
{
   
   public static String GetLogFileName_s()
   {
      return LogFileName_s;
   }
   
   public static String GetVersion_s()
   {
      return Version_s;
   }

   public static void main(final String[] args)
   {
      if ( args.length >= 1 )
      {
         File newFile = new File(args[0]);
         if ( newFile.isFile() )
         {
            String files[] = new String[1];
            files[0] = newFile.getAbsolutePath();
            LogFileName_s = files[0] + ".log";
         }
      }
      Log.Init(new File(LogFileName_s));
      Log.info("Starting ...");
      
      // Get the version string from the jar file
      try
      {
         Version_s = Main.class.getPackage().getImplementationVersion();
         if ( Version_s == null )
         {
            throw new Exception();
         }
      }
      catch ( Exception e )
      {
         Version_s = "not set";
      }
      Log.info("version : " + Version_s);

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
   
   // --- PRIVATE

   private static String LogFileName_s = "ExcelToolkit.log";
   private static String Version_s = "";

}
