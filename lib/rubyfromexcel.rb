# Standard library
require 'fileutils'
require 'singleton'

# Gems
require 'nokogiri'
require 'rubyscriptwriter'
require 'rubyspecwriter'
require 'rubypeg'

# Components of RubyFromExcel
require_relative 'excelfile/excelfile'
require_relative 'formulae/formulae'
require_relative 'cells/cells'
require_relative 'optimiser/optimiser'

require_relative 'runtime/runtime_formula_builder'

module RubyFromExcel
  class Process
    attr_accessor :source_excel_directory
    attr_accessor :target_ruby_directory
    attr_accessor :workbook
    attr_accessor :skip_tests
    attr_accessor :prune_except_output_sheets
    attr_accessor :convert_independent_of_input_sheets
  
    def initialize(&block)
      instance_eval(&block) if block
    end
  
    def workbook_filename
      File.join(source_excel_directory,'xl','workbook.xml')
    end
  
    def start!
      reset_global_classes
      
      time "Preparing destination folder..." do
        prepare_destination_folder
      end
    
      time "Loading..." do
        self.workbook = Workbook.new(workbook_filename)
      end
    
      if prune_except_output_sheets
        time "Pruning..." do
          workbook.prune_cells_not_needed_for_output_sheets(*prune_except_output_sheets)
          workbook.convert_cells_to_values_when_independent_of_input_sheets(*convert_independent_of_input_sheets)
        end
      end
        
      time "Workbook contains #{workbook.worksheets.size} sheets:\n" do
    
        puts "0) Generating ruby for the workbook"
        write workbook.to_ruby, :to, "spreadsheet.rb"
    
        i = 0
        workbook.worksheets.each do |variable_name, worksheet|
          time "#{i+=1}) Generating ruby for #{variable_name}..." do
            print 'ruby...'
            write worksheet.to_ruby, :to, 'sheets', "#{variable_name}.rb"
            unless skip_tests
              print 'test...'
              write worksheet.to_test, :to, 'specs',"#{variable_name}_rspec.rb"
            end
            # worksheet.nil_memory_consuming_variables!
          end
        end
    
      end
      unless skip_tests
        puts "Running tests of generated files"
        puts `spec -fo #{File.join(target_ruby_directory,'specs',"*")}`
      end
    end
    
    # FIXME: Urgh. Global variables. Need to eliminate these!
    def reset_global_classes
      SheetNames.instance.clear
      SharedStrings.instance.clear
    end
    
    def prepare_destination_folder
      FileUtils.mkpath(File.join(target_ruby_directory,'specs'))
      FileUtils.mkpath(File.join(target_ruby_directory,'sheets'))
    end
  
    def write(thing,to,*filenames)
      File.open(File.join(target_ruby_directory,*filenames),'w') do |f|
        f.puts thing.to_s
      end
    end
  
    def time(message)
      print message
      STDOUT.flush
      start_time = Time.now
      yield
      puts "took #{(Time.now-start_time).to_i} seconds."
    end
  
  end
end