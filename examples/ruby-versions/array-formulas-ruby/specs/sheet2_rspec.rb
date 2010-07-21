# coding: utf-8
require_relative '../spreadsheet'
# Sheet2
describe 'Sheet2' do
  def sheet2; $spreadsheet ||= Spreadsheet.new; $spreadsheet.sheet2; end

end

