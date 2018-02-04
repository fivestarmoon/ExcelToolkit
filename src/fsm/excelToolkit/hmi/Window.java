package fsm.excelToolkit.hmi
;
import java.awt.Container;
import java.awt.Dimension;
import java.awt.Point;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.dnd.DnDConstants;
import java.awt.dnd.DropTarget;
import java.awt.dnd.DropTargetDragEvent;
import java.awt.dnd.DropTargetDropEvent;
import java.awt.dnd.DropTargetEvent;
import java.awt.dnd.DropTargetListener;
import java.awt.event.*;
import javax.swing.*;

import fsm.common.Log;
import fsm.common.parameters.Parameters;
import fsm.excelToolkit.hmi.table.TableSpreadsheet;
import fsm.excelToolkit.wpsr.WpsrSummaryPanel;

@SuppressWarnings("serial")
public class Window extends JFrame
implements WindowListener, DropTargetListener
{
   public void createAndShowGUI()
   {
      // Create and set up the window.
      setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

      // Set the main window frame's title
      setTitle("Excel Toolkit");
      setResizable(true);

      // Set the window size
      setSize(new Dimension(700,700));
      setLocation(new Point(100,100));
      addWindowListener(this);

      // Create the menu
      JMenuBar menuBar = new JMenuBar();
      // File menu
      JMenu fileMenu = new JMenu("File");
      fileMenu.setMnemonic(KeyEvent.VK_F);
      fileMenu.setMnemonic(KeyEvent.VK_F);
      JMenuItem openItem = new JMenuItem("Open ...");
      openItem.setMnemonic(KeyEvent.VK_O);
      openItem.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_O, InputEvent.CTRL_MASK));
      openItem.addActionListener(new ActionListener()
         {
            @Override
            public void actionPerformed(ActionEvent e)
            {
               JFileChooser chooser = new JFileChooser(".");
               int returnVal = chooser.showOpenDialog(Window.this);
               if(returnVal == JFileChooser.APPROVE_OPTION)
               {
                  processNewParameterFile(chooser.getSelectedFile().getAbsolutePath());
               }
            }         
         });
      fileMenu.add(openItem);
      JMenuItem quitItem = new JMenuItem("Exit");
      quitItem.setMnemonic(KeyEvent.VK_X);
      quitItem.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_X, InputEvent.CTRL_MASK));
      quitItem.addActionListener(new ActionListener()
      {
         @Override
         public void actionPerformed(ActionEvent e)
         {
            Window.this.dispatchEvent(new WindowEvent(Window.this, WindowEvent.WINDOW_CLOSING));
         }      
      });
      fileMenu.add(quitItem);
      menuBar.add(fileMenu);
      setJMenuBar(menuBar);

      // Display the window.
      showContent(new JLabel("Nothing to display ..."));
      setVisible(true);
   }
   
   public void showContent(Container content)
   {
      setContentPane(content);
      new DropTarget(content, this);
      revalidate(); 
      repaint();
   }

   public void processNewParameterFile(String absolutePath)
   {   
      try
      {
         params_ = new Parameters(absolutePath);
         String type = params_.getReader().getStringValue("type", "");
         if ( "wpsr_summary".equalsIgnoreCase(type) )
         {
            showContent(new JLabel("Loading WPSR summary ..."));
            table_ = new WpsrSummaryPanel();
            table_.createTable(this, params_); // eventually calls showContent         
         }
         else
         {
            showContent(new JLabel("Did not recogonize json \"type\" [" + type + "]"));
         }
      }
      catch (Exception e)
      {
         showContent(new JLabel("Error in json " + absolutePath));
         return;
      }

   }

   @Override
   public void windowOpened(WindowEvent e)
   {      
   }

   @Override
   public void windowClosing(WindowEvent e)
   {
      Log.info("closing window ...");
      
   }

   @Override
   public void windowClosed(WindowEvent e)
   {
      Log.info("done.");
      
   }

   @Override
   public void windowIconified(WindowEvent e)
   {
   }

   @Override
   public void windowDeiconified(WindowEvent e)
   {
   }

   @Override
   public void windowActivated(WindowEvent e)
   {
   }

   @Override
   public void windowDeactivated(WindowEvent e)
   {
   }
   
   @Override
   public void dragEnter(DropTargetDragEvent dtde)
   {
   }

   @Override
   public void dragOver(DropTargetDragEvent dtde)
   {
   }

   @Override
   public void dropActionChanged(DropTargetDragEvent dtde)
   {
   }

   @Override
   public void dragExit(DropTargetEvent dte)
   {
   }

   @Override
   public void drop(DropTargetDropEvent dtde)
   {
      try
      {
         // Ok, get the dropped object and try to figure out what it is
         Transferable tr = dtde.getTransferable();
         DataFlavor[] flavors = tr.getTransferDataFlavors();
         for (int i = 0; i < flavors.length; i++)
         {
            System.out.println("Possible flavor: " + flavors[i].getMimeType());
            // Check for file lists specifically
            if (flavors[i].isFlavorJavaFileListType())
            {
               // Great!  Accept copy drops...
               dtde.acceptDrop(DnDConstants.ACTION_COPY_OR_MOVE);
               System.out.println("Successful file list drop.");

               // And add the list of file names to our text area
               @SuppressWarnings("unchecked")
               java.util.List<Object> list = (java.util.List<Object>)tr.getTransferData(flavors[i]);
               if ( list.size() >= 1)
               {
                  String files[] = new String[list.size()];
                  for ( int ii=0; ii<files.length; ii++ )
                  {
                     files[ii] = list.get(ii).toString();
                  }
                  processNewParameterFile(files[0]);
                  dtde.dropComplete(true);
                  return;
               }
            }
         }
         // User must not have dropped a file list
         Log.info("Drop failed: " + dtde);
         dtde.rejectDrop();
      }
      catch (Exception e)
      {
         Log.severe("Drop exception: ", e);
      }
      
   }
   
   // PROTECTED
   
   // PRIVATE
   
   private Parameters params_;
   private TableSpreadsheet table_;
}
