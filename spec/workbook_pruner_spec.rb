require_relative 'spec_helper'

describe WorkbookPruner do

  it "should initialize with a workbook and ensure that all the worksheets in that workbook check their dependencies" do
    workbook = mock(:workbook)
    WorkbookPruner.new(workbook) 
  end
  
  it "should be able prune any cells not required" do
    workbook = mock(:workbook)
    workbook.should_receive(:work_out_dependencies)

    sheet1 = mock(:worksheet,:name =>'sheet1')
    sheet2 = mock(:worksheet,:name =>'sheet2')

    workbook.should_receive(:worksheets).at_least(:once).and_return({'sheet1' => sheet1,'sheet2' => sheet2})
    workbook.should_receive(:total_cells).and_return(3)

    cell1 = mock(:cell,:worksheet => sheet1,:reference => 'a1')
    #workbook.should_receive(:cell).with('sheet1.a1').and_return(cell1)
    cell1.should_receive(:dependencies).and_return(['sheet2.a2'])

    cell2 = mock(:cell,:worksheet => sheet2,:reference => 'a2')
    workbook.should_receive(:cell).with('sheet2.a2').and_return(cell2)
    cell2.should_receive(:dependencies).and_return([])

    cell3 = mock(:cell,:worksheet => sheet2,:reference => 'a3')
    #workbook.should_receive(:cell).with('sheet2.a3').and_return(cell3)
    #cell3.should_receive(:dependencies).and_return(['sheet1.a1'])

    sheet1_cells = {'a1' => cell1}
    sheet2_cells = {'a2' => cell2,'a3' => cell3}
    sheet1.should_receive(:cells).at_least(:once).and_return(sheet1_cells)
    sheet2.should_receive(:cells).at_least(:once).and_return(sheet2_cells)

    SheetNames.instance.clear
    SheetNames.instance['Output Sheet'] = 'sheet1'
    
    wb = WorkbookPruner.new(workbook)
    workbook.work_out_dependencies
    wb.prune_cells_not_needed_for_output_sheets('Output Sheet')

    sheet1_cells.should have_key('a1')
    sheet2_cells.should have_key('a2')
    sheet2_cells.should_not have_key('a3')
  end
end