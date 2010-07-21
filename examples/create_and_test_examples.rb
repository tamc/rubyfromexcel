require_relative '../lib/rubyfromexcel'

def convert(basename)
    # Need the original spreadsheet
    spreadsheet = File.join(File.dirname(__FILE__),'sheets',"#{basename}.xlsx")
    
    # A place to put an unzipped version of the spreadsheet (could be a tmp dir, but helpful for debugging if local)
    unzipped_spreadsheet = File.join(File.dirname(__FILE__),'unzipped-sheets',basename)
    
    # A place to put the resulting ruby version
    ruby_version = File.join(File.dirname(__FILE__),'ruby-versions',"#{basename}-ruby")

    puts "Converting #{spreadsheet} into #{ruby_version}"

    # The spreadsheet needs to be unzipped before starting
    puts `unzip -uo #{spreadsheet} -d #{unzipped_spreadsheet}`

    RubyFromExcel::Process.new do
      self.source_excel_directory = unzipped_spreadsheet
      self.target_ruby_directory = ruby_version
      self.skip_tests = false
      if basename == "pruning"
        self.prune_except_output_sheets = ['Outputs']
        self.convert_independent_of_input_sheets = ['Inputs']
      end
    end.start!

    puts
end

if ARGV[0]
  convert ARGV[0]
else
  %w{array-formulas complex-test namedReferenceTest sharedFormulaTest table-test pruning}.each do |basename|
    convert basename
  end
end