package fsm.excelToolkit;

import java.io.File;

import fsm.common.Log;
import fsm.excelToolkit.hmi.Window;

public class Main
{

   public static void main(final String[] args)
   {
      Log.Init();
      Log.info("Starting ...");

      
      // Schedule a job for the event-dispatching thread:
        // creating and showing this application's GUI.
        javax.swing.SwingUtilities.invokeLater(new Runnable()
        {
            public void run()
            {
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