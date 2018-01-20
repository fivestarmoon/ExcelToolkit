package fsm.excelToolkit.hmi;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;

import javax.swing.BorderFactory;
import javax.swing.JLabel;
import javax.swing.JPanel;

@SuppressWarnings("serial")
public class TableSpreadsheet extends JPanel
{
   public TableSpreadsheet()
   {
      l_ = new GridBagLayout();
      setLayout(l_);
      reset();
      this.setMinimumSize(new Dimension(400,400));
   }
   
   public void reset()
   {
      c_ = new GridBagConstraints();
      c_.gridx = 0;
      c_.gridy = 0;
      resetFormatting();
      removeAll();
   }
   
   public void addRow()
   {
      c_.gridx = 0;
      c_.gridy++;
   }
   
   public void addLabel(JLabel label, int width, boolean greedy, boolean centered)
   {
      label.setBorder(BorderFactory.createLineBorder(Color.black));
      c_.gridwidth = width;
      if ( greedy ) c_.weightx = 1.0;
      if ( centered ) c_.anchor = GridBagConstraints.CENTER;
      add(label, c_);
      resetFormatting();
   }
   
   public void addCell(JLabel label, int width, boolean greedy, boolean centered)
   {
      label.setBorder(BorderFactory.createLineBorder(Color.black));
      c_.gridwidth = width;
      if ( greedy ) c_.weightx = 1.0;
      if ( centered ) c_.anchor = GridBagConstraints.CENTER;
      add(label, c_);
      resetFormatting();
   }
   
   // --- PRIVATE
   
   private void resetFormatting()
   {
      c_.gridwidth = 1; c_.ipadx = 0; c_.weightx = 0.0; 
      c_.gridheight = 1; c_.ipady = 0; c_.weighty = 0.0;
      c_.fill = GridBagConstraints.HORIZONTAL; // VERTICAL BOTH
      c_.insets = new Insets(0,0,0,0);
      c_.anchor = GridBagConstraints.LINE_START;
      // FIRST_LINE_START  PAGE_START  FIRST_LINE_END
      // LINE_START        CENTER      LINE_END
      // LAST_LINE_START   PAGE_END    LAST_LINE_END      
   }
   
   private GridBagLayout l_;
   private GridBagConstraints c_;

}
