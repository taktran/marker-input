#!/usr/bin/env ruby
require 'fileutils'
require 'optparse'
require 'docx'

OPTIONS = {}

optparse = OptionParser.new do|opts|
  opts.banner = <<-eos
Import marks from .docx files

Usage: #{__FILE__} [options] [file or folder]
  eos

  opts.on( '-h', '--help', 'Display this screen' ) do
    puts opts
    exit
  end
end

optparse.parse!

INPUT = ARGV[0]
STUDENT_ID_REGEX = /Student ID: *(\d*)/i
MARK_REGEX = /Mark: *(\d*)/i

def parse_file(file)
  assessment = Docx::Document.open(file)
  student_id = nil
  mark = nil
  assessment.each_paragraph do |p|
    p_string = p.to_s

    student_id_match = STUDENT_ID_REGEX.match(p_string)
    student_id = student_id_match[1] if student_id_match

    mark_match = MARK_REGEX.match(p_string)
    mark = mark_match[1] if mark_match
  end

  puts "Student ID = #{student_id ? student_id : '[not found]'}"
  puts "Mark = #{mark ? mark : '[not found]'}"
end

# Handle file or folder input
if INPUT.nil?
  puts "No file specified"
else
  if File.exists? INPUT
    if File.directory? INPUT
      Dir.entries(INPUT).each do |file|
        file_path = "#{INPUT}#{"/" unless INPUT[0].chr == '/'}#{file}"
        parse_file(file_path) unless File.directory?(file_path)
      end
    else
      parse_file INPUT
    end
  else
    puts "Invalid file/folder"
  end
end