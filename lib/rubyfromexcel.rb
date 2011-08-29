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
  
  def self.debug(name,message)
    return "" unless $DEBUG == true
    file_name = "#{name.to_s}-debug.txt"
    @files ||= {}
    file = @files[file_name] ||= File.open(file_name.to_s,'w')
    file.puts message
  end
  
  class Process
    attr_accessor :source_excel_directory
    attr_accessor :target_ruby_directory
    attr_accessor :workbook
    attr_accessor :skip_tests
    attr_accessor :prune_except_output_sheets
    attr_accessor :convert_independent_of_input_sheets
    attr_accessor :checkpoint_directory
    attr_accessor :stage
    attr_accessor :debug_dont_write_checkpoint_after_stage
  
    def initialize(&block)
      instance_eval(&block) if block
      load_from_checkpoint
    end
  
    def workbook_filename
      File.join(source_excel_directory,'xl','workbook.xml')
    end
  
    def start!(starting_stage = nil)
      self.stage = starting_stage || self.stage || 0
      
      checkpoint 0 do
        reset_global_classes
      
        time "Preparing destination folder..." do
          prepare_destination_folder
        end
    
        time "Loading..." do
          self.workbook = Workbook.new(workbook_filename)
        end
      end

      time "Pruning..." do    
        
        if prune_except_output_sheets || convert_independent_of_input_sheets
          workbook.work_out_dependencies
        end
        
        checkpoint 1 do
          if convert_independent_of_input_sheets
            workbook.convert_cells_to_values_when_independent_of_input_sheets(*convert_independent_of_input_sheets)
          end
        end
        
        checkpoint 2 do
          if prune_except_output_sheets
            workbook.prune_cells_not_needed_for_output_sheets(*prune_except_output_sheets)
          end
        end
        
      end
        
      time "Workbook contains #{workbook.worksheets.size} sheets:\n" do
        
        checkpoint 3 do
          puts "0) Generating ruby for the workbook"
          write "spreadsheet.rb" do
            workbook.to_ruby
          end
        end
    
        checkpoint 4 do
          i = 0
          workbook.worksheets.each do |variable_name, worksheet|
            time "#{i+=1}) Generating ruby for #{variable_name}..." do
              write 'sheets', "#{variable_name}.rb" do
                worksheet.to_ruby
              end
            end
          end
        end

        checkpoint 5 do
          unless skip_tests
            i = 0
            workbook.worksheets.each do |variable_name, worksheet|
              time "#{i+=1}) Generating spec for #{variable_name}..." do
                write 'specs',"#{variable_name}_rspec.rb" do
                  worksheet.to_test
                end
              end
            end
          end # skip tests
        end # 5 
        
      end
      
      unless skip_tests
        puts "Running tests of generated files"
        puts `rspec -fp #{File.join(target_ruby_directory,'specs',"*")}`
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
  
    def write(*filenames)
      target_filename = File.join(target_ruby_directory,*filenames)
      if checkpoint_directory && !debug_dont_write_checkpoint_after_stage && File.exists?(target_filename)
        print "skipping #{File.basename(target_filename)}, already exists and checkpoints are on..."
        return nil
      end
      File.open(target_filename,'w') do |f|
        f.puts yield.to_s
      end
    end
    
    def checkpoint(checkpoint_number)
      puts
      if self.stage > checkpoint_number
        puts "Stage #{checkpoint_number} already completed, skipping..."
        return
      else
        puts "Stage #{checkpoint_number}"
      end
      yield
      save_checkpoint if checkpoint_directory  # Then create a check point
      self.stage = checkpoint_number + 1
    end
    
    def save_checkpoint
      if debug_dont_write_checkpoint_after_stage && (self.stage > debug_dont_write_checkpoint_after_stage)
        puts "Debug mode: Not writing checkpoint for stage #{checkpoint_number}"
        return
      else
        puts "Writing checkpoint for stage #{self.stage}"
        FileUtils.mkpath(checkpoint_directory)
        objects_to_dump = [self.stage,self.workbook,SheetNames.instance.to_hash,SharedStrings.instance.to_a]
        File.open(checkpoint_filename,'w') { |f| f.puts Marshal.dump(objects_to_dump) }
      end
    end

    def load_from_checkpoint
      unless checkpoint_directory
        puts "No checkpoint directory given"
        return false
      end
      checkpoint_filenames = Dir.entries(checkpoint_directory)
      highest_stage_checkpoint = checkpoint_filenames.map { |filename| filename =~ /checkpoint(\d+)/ ? $1.to_i : 0 }.max
      checkpoint_to_open = File.join(checkpoint_directory,"checkpoint#{highest_stage_checkpoint}.marshal")
      unless File.exists?(checkpoint_to_open)
        puts "No checkpoint file found"
        return false
      end
      dumped_objects = Marshal.load(File.open(checkpoint_to_open))
      self.stage = dumped_objects.shift + 1
      self.workbook = dumped_objects.shift
      SheetNames.instance.replace(dumped_objects.shift)
      SharedStrings.instance.replace(dumped_objects.shift)
      puts "Stage #{self.stage-1} checkpoint found and loaded."
    end
    
    def checkpoint_filename
      File.join(checkpoint_directory,"checkpoint#{self.stage}.marshal")
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