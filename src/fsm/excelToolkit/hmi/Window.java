package fsm.excelToolkit.hmi
;
import java.awt.Container;
import java.awt.Desktop;
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
import java.io.File;
import java.io.IOException;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

import fsm.common.FsmResources;
import fsm.common.Log;
import fsm.common.parameters.Parameters;
import fsm.common.utils.FileModifiedListener;
import fsm.common.utils.FileModifiedMonitor;
import fsm.excelToolkit.Main;
import fsm.excelToolkit.generic.GenericSheetPanel;
import fsm.excelToolkit.hmi.table.TableSpreadsheet;
import fsm.excelToolkit.jira.JiraSummaryPanel;
import fsm.excelToolkit.loading.LoadingSummaryPanel;
import fsm.excelToolkit.wpsr.WpsrSummaryPanel;

@SuppressWarnings("serial")
public class Window extends JFrame
implements WindowListener, DropTargetListener, FileModifiedListener
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
      
      // File > Open
      {
         JMenuItem menuItem = new JMenuItem(
            "Open JSON file ...", 
            FsmResources.getIconResource("file_empty.png"));
         menuItem.setMnemonic(KeyEvent.VK_O);
         menuItem.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_O, InputEvent.CTRL_MASK));
         menuItem.addActionListener(new ActionListener()
         {
            @Override
            public void actionPerformed(ActionEvent e)
            {
               JFileChooser chooser = new JFileChooser(".");
               chooser.addChoosableFileFilter(new FileNameExtensionFilter("Excel tool kit files", "extk"));
               chooser.addChoosableFileFilter(new FileNameExtensionFilter("Text files", "txt"));
               int returnVal = chooser.showOpenDialog(Window.this);
               if(returnVal == JFileChooser.APPROVE_OPTION)
               {
                  processNewParameterFile(chooser.getSelectedFile().getAbsolutePath());
               }
            }         
         });
         fileMenu.add(menuItem);
      }
      
      // File > Reload
      {
         JMenuItem menuItem = new JMenuItem(
            "Reload JSON file", 
            FsmResources.getIconResource("program_refresh.png"));
         menuItem.setMnemonic(KeyEvent.VK_R);
         menuItem.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_R, InputEvent.CTRL_MASK));
         menuItem.addActionListener(new ActionListener()
         {
            @Override
            public void actionPerformed(ActionEvent e)
            {
               if ( lastFile_.length() > 0 )
               {
                  processNewParameterFile(lastFile_);
               }
            }         
         });
         fileMenu.add(menuItem);
      }
      
      // File > Edit
      {
         JMenuItem menuItem = new JMenuItem(
            "Edit JSON file ...", 
            FsmResources.getIconResource("file_paste.png"));
         menuItem.setMnemonic(KeyEvent.VK_E);
         menuItem.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_E, InputEvent.CTRL_MASK));
         menuItem.addActionListener(new ActionListener()
         {
            @Override
            public void actionPerformed(ActionEvent e)
            {
               try
               {
                  if ( lastFile_.length() > 0 )
                  {
                     Desktop.getDesktop().open(new File(lastFile_));
                  }
               }
               catch (IOException e1)
               {
                  Log.severe("Could not open file with desktop", e1);
               }
            }         
         });
         fileMenu.add(menuItem);
      }
      
      // File > Exit
      {
         JMenuItem menuItem = new JMenuItem(
            "Exit", 
            FsmResources.getIconResource("program_exit.png"));
         menuItem.setMnemonic(KeyEvent.VK_X);
         menuItem.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_X, InputEvent.CTRL_MASK));
         menuItem.addActionListener(new ActionListener()
         {
            @Override
            public void actionPerformed(ActionEvent e)
            {
               Window.this.dispatchEvent(new WindowEvent(Window.this, WindowEvent.WINDOW_CLOSING));
            }      
         });
         fileMenu.add(menuItem);
      }
      
      // Help menu
      JMenu helpMenu = new JMenu("Help");
      helpMenu.setMnemonic(KeyEvent.VK_H);
      
      // Help > View Log
      {
         JMenuItem menuItem = new JMenuItem(
            "View log ...");
         menuItem.addActionListener(new ActionListener()
         {
            @Override
            public void actionPerformed(ActionEvent e)
            {
               try
               {
                  Desktop.getDesktop().open(new File(Main.GetLogFileName_s()));
               }
               catch (IOException e1)
               {
                  Log.severe("Could not open file with desktop", e1);
               }
            }      
         });
         helpMenu.add(menuItem);
      }      
      
      // Help > About
      {
         JMenuItem menuItem = new JMenuItem(
            "About");
         menuItem.addActionListener(new ActionListener()
         {
            @Override
            public void actionPerformed(ActionEvent e)
            {
               JOptionPane.showMessageDialog(Window.this, "Excel Toolkit\nKurt Hagen\n" + Main.GetBuildDate_s());
            }      
         });
         helpMenu.add(menuItem);
      }      
      
      // Add menus to the bar
      menuBar.add(fileMenu);
      menuBar.add(Box.createHorizontalGlue());
      menuBar.add(helpMenu);
      
      // Add menu bar
      setJMenuBar(menuBar);

      // Display the window.
      title_ = "";
      showContent(new JLabel("Nothing to display ..."));
      setVisible(true);
   }
   
   public void showContent(Container content)
   {
      setTitle("Excel Toolkit [" + title_ + "]");
      setContentPane(content);
      new DropTarget(content, this);
      revalidate(); 
      repaint();
   }

   public void processNewParameterFile(String absolutePath)
   {   
      try
      {
         // Update the working directory
         File file = new File(absolutePath);
         String path = file.getParent();
         if ( path != null )
         {
            System.setProperty("user.dir", path);
         }
         
         if ( fileModifiedMonitor_ != null )
         {
            fileModifiedMonitor_.stop();
            fileModifiedMonitor_ = null;
         }         
         lastFile_ = absolutePath;
         params_ = new Parameters(absolutePath);
         if ( table_ != null )
         {
            table_.destroyTable();
         }
         String type = params_.getReader().getStringValue("type", "");
         if ( "wpsr_summary".equalsIgnoreCase(type) )
         {
            setSize(new Dimension(
               (int) params_.getReader().getLongValue("width",  700),
               (int) params_.getReader().getLongValue("height",  700)));
            title_ = new File(absolutePath).getName();
            showContent(new JLabel("Loading WPSR summary ..."));
            table_ = new WpsrSummaryPanel();
            table_.createTable(this, params_); // eventually calls showContent         
         }
         else if ( "loading".equalsIgnoreCase(type) )
         {
            setSize(new Dimension(
               (int) params_.getReader().getLongValue("width",  700),
               (int) params_.getReader().getLongValue("height",  700)));
            title_ = new File(absolutePath).getName();
            showContent(new JLabel("Loading spreadsheet loading ..."));
            table_ = new LoadingSummaryPanel();
            table_.createTable(this, params_); // eventually calls showContent         
         }
         else if ( "jira_summary".equalsIgnoreCase(type) )
         {
            setSize(new Dimension(
               (int) params_.getReader().getLongValue("width",  700),
               (int) params_.getReader().getLongValue("height",  700)));
            title_ = new File(absolutePath).getName();
            showContent(new JLabel("Loading JIRA summary ..."));
            table_ = new JiraSummaryPanel();
            table_.createTable(this, params_); // eventually calls showContent         
         }
         else if ( "spreadsheets".equalsIgnoreCase(type) )
         {
            setSize(new Dimension(
               (int) params_.getReader().getLongValue("width",  700),
               (int) params_.getReader().getLongValue("height",  700)));
            title_ = new File(absolutePath).getName();
            showContent(new JLabel("Loading generic spreadsheets ..."));
            table_ = new GenericSheetPanel();
            table_.createTable(this, params_); // eventually calls showContent         
         }
         else
         {
            showContent(new JLabel("Did not recogonize json \"type\" [" + type + "]"));
         }
      }
      catch (Exception e)
      {
         showContent(new JLabel("Error in json " + absolutePath + ". See log for more information."));
         return;
      }
      fileModifiedMonitor_ = new FileModifiedMonitor(new File(absolutePath), this);
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

   @Override
   public void fileModified()
   {
      SwingUtilities.invokeLater(
         new Runnable()
         {
            public void run()
            {
               processNewParameterFile(lastFile_);
            }
         });
   }
   
   // PROTECTED
   
   // PRIVATE
   
   private Parameters params_;
   private TableSpreadsheet table_;
   private String title_ = "";
   private String lastFile_ = "";
   private FileModifiedMonitor fileModifiedMonitor_ = null;
}
