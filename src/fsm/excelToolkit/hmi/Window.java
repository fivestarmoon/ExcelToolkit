package fsm.excelToolkit.hmi
;
import java.awt.Dimension;
import java.awt.Point;
import java.awt.event.*;
import javax.swing.*;

import fsm.common.Log;
import fsm.common.parameters.Parameters;
import fsm.common.parameters.Reader;
import fsm.spreadsheet.SsFile;
import fsm.spreadsheet.SsTable;

@SuppressWarnings("serial")
public class Window extends JFrame
implements ActionListener, WindowListener
{
   public void createAndShowGUI()
   {
      // Create and set up the window.
      setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

      // Set the main window frame's title
      setTitle("Excel Toolkit");
//      ArrayList<Image> icons = new ArrayList<Image>();
//      icons.add(LvApplication.getImageResource("icon16x16.png"));
//      icons.add(LvApplication.getImageResource("icon32x32.png"));
//      setIconImages(icons);
      setResizable(true);

      // Add the main window components
      //updateLayout();

      // Set the window size
      setSize(new Dimension(400,400));
      setLocation(new Point(400,400));
      this.addWindowListener(this);

      // Display the window.
      //pack();
      //new DropTarget(viewer, this);
      //new DropTarget(viewerSplit, this);
      
      ss_ = new TableSpreadsheet();
      setVisible(true);
      ss_.setOpaque(true); //content panes must be opaque
      setContentPane(new JScrollPane(ss_));

      //Display the window.
      setVisible(true);
   }

   public void processNewParameterFile(String absolutePath)
   {
      
      Parameters params = new Parameters(absolutePath);
      Reader reader = params.getReader();
      SsFile file = SsFile.Create(reader.getStringValue("file", ""));
      SsTable table = new SsTable(file, (int)reader.getLongValue("sheetOffset", 0));
      table.addRow(-1 + (int)reader.getLongValue("startRow", 0), 
                   -1 + (int)reader.getLongValue("endRow", 0));
      table.addCol(SsTable.ColumnLettersToIndex(reader.getStringValue("resourceCol", "")));
      table.addCol(SsTable.ColumnLettersToIndex(reader.getStringValue("budgetCol", "")));
      table.addCol(SsTable.ColumnLettersToIndex(reader.getStringValue("prevActualCol", "")));
      table.addCol(SsTable.ColumnLettersToIndex(reader.getStringValue("thisActualCol", "")));
      table.addCol(SsTable.ColumnLettersToIndex(reader.getStringValue("etcCol", "")));
      table.readTable();
      Log.info(table.toString());
      for ( int row : table.getRowIterator() )
      {
         ss_.addRow();
         ss_.addLabel(new JLabel(table.getCell(row, 0).toString()), 1, true, false);
         for ( int col : table.getColIterator() )
         {
            Log.info(String.format("[%d][%d]=%s", row, col, table.getCell(row, col)));             
         }
      }
   }

   @Override
   public void windowOpened(WindowEvent e)
   {
      // TODO Auto-generated method stub
      
   }

   @Override
   public void windowClosing(WindowEvent e)
   {
      // TODO Auto-generated method stub
      Log.info("closing window ...");
      
   }

   @Override
   public void windowClosed(WindowEvent e)
   {
      // TODO Auto-generated method stub
      Log.info("done.");
      
   }

   @Override
   public void windowIconified(WindowEvent e)
   {
      // TODO Auto-generated method stub
      
   }

   @Override
   public void windowDeiconified(WindowEvent e)
   {
      // TODO Auto-generated method stub
      
   }

   @Override
   public void windowActivated(WindowEvent e)
   {
      // TODO Auto-generated method stub
      
   }

   @Override
   public void windowDeactivated(WindowEvent e)
   {
      // TODO Auto-generated method stub
      
   }

   @Override
   public void actionPerformed(ActionEvent e)
   {
      // TODO Auto-generated method stub
      
   }
   
   // PRIVATE
   
   private TableSpreadsheet ss_;
}
