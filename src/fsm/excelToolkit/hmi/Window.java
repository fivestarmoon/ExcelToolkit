package fsm.excelToolkit.hmi
;
import java.awt.Dimension;
import java.awt.Point;
import java.awt.event.*;
import javax.swing.*;

import fsm.common.Log;
import fsm.common.parameters.Parameters;
import fsm.excelToolkit.hmi.table.TableSpreadsheet;

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
      setSize(new Dimension(600,500));
      setLocation(new Point(400,400));
      this.addWindowListener(this);

      // Display the window.
      //pack();
      //new DropTarget(viewer, this);
      setContentPane(new JLabel("Nothing to display ..."));

      //Display the window.
      setVisible(true);
   }

   public void processNewParameterFile(String absolutePath)
   {      
      params_ = new Parameters(absolutePath);
      setContentPane(new JLabel("Nothing to display ..."));
      table_ = new WpsrSummaryPanel();
      table_.createTable(this, params_);
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
   
   // PROTECTED
   
   public void showTable(TableSpreadsheet table)
   {
      setContentPane(table);
      revalidate(); 
      repaint();
   }
   
   // PRIVATE
   
   private Parameters params_;
   private TableSpreadsheet table_;
}
