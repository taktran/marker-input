#!/usr/bin/env ruby
require 'docx'

student_id_regex = /Student ID: *(\d*)/i
mark_regex = /Mark: *(\d*)/i

assessment = Docx::Document.open('docs/assessment.docx')
student_id = nil
mark = nil
assessment.each_paragraph do |p|
  p_string = p.to_s

  student_id_match = student_id_regex.match(p_string)
  student_id = student_id_match[1] if student_id_match

  mark_match = mark_regex.match(p_string)
  mark = mark_match[1] if mark_match
end

puts "Student ID = #{student_id ? student_id : '[not found]'}"
puts "Mark = #{mark ? mark : '[not found]'}"