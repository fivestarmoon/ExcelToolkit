package fsm.excelToolkit.hmi;

import fsm.common.Log;
import fsm.common.parameters.Reader;
import fsm.excelToolkit.hmi.table.TableSpreadsheet;
import fsm.spreadsheet.SsFile;
import fsm.spreadsheet.SsTable;

@SuppressWarnings("serial")
public class WpsrSummaryPanel extends TableSpreadsheet
{
   public WpsrSummaryPanel()
   {
   }
   
   @Override
   public void createPanel()
   {      
      setColumns(
                 new String[] {"Resource", "Budget", "Prev Actual", "Week Actual", "ETC" },
                 false);
      
      Reader reader = getParameters().getReader();
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
      startAddRows();
      for ( int row : table.getRowIterator() )
      {
         addRowOfCells(table.getCellsForRow(row));
         for ( int col : table.getColIterator() )
         {
            Log.info(String.format("[%d][%d]=%s", row, col, table.getCell(row, col)));             
         }
      }
      stopAddRows();
      
      // Display the panel
      displayPanel();
   }
   
   // --- PRIVATE

}
