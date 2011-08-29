require_relative '../lib/rubyfromexcel'

$DEBUG = false

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
      case basename
      when "pruning"
        self.prune_except_output_sheets = ['Outputs']
        self.convert_independent_of_input_sheets = ['Inputs']      
      when "checkpoint"
        self.checkpoint_directory =  File.join(File.dirname(__FILE__),'checkpoints','checkpoint')
        # self.debug_dont_write_checkpoint_after_stage = 1
      when "2050Model", "2050ModelCutDown"
        self.prune_except_output_sheets = ['Intermediate output','Control']
        self.convert_independent_of_input_sheets = ['Control']
        # self.checkpoint_directory =  File.join(File.dirname(__FILE__),'checkpoints','2050Model')        
      end
    end.start!

    puts
end

if ARGV[0]
  convert ARGV[0]
else
  %w{array-formulas complex-test namedReferenceTest sharedFormulaTest table-test pruning checkpoint}.each do |basename|
    convert basename
  end
end