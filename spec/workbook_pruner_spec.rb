require_relative 'spec_helper'

describe WorkbookPruner do

  it "should initialize with a workbook and ensure that all the worksheets in that workbook check their dependencies" do
    workbook = mock(:workbook)
    WorkbookPruner.new(workbook) 
  end
  
  it "should be able prune any cells not required" do
    workbook = mock(:workbook)
    workbook.should_receive(:work_out_dependencies)
    sheet1 = mock(:worksheet)
    workbook.should_receive(:worksheets).at_least(:once).and_return({'sheet1' => sheet1})
    workbook.should_receive(:total_cells).and_return(2)
    cell1 = mock(:cell)
    sheet1.should_receive(:cells).at_least(:once).and_return({'a1' => cell1})
    cell1.should_receive(:dependencies).and_return(['sheet1.a2'])
    cell2 = mock(:cell)
    workbook.should_receive(:cell).with('sheet1.a2').and_return(cell2)
    cell2.should_receive(:dependencies).and_return([])
    SheetNames.instance.clear
    SheetNames.instance['Output Sheet'] = 'sheet1'
    wb = WorkbookPruner.new(workbook)
    wb.prune_cells_not_needed_for_output_sheets('Output Sheet')
  end
end