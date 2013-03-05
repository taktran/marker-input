#!/usr/bin/env ruby
require 'docx'

assessment = Docx::Document.open('docs/assessment.docx')
assessment.each_paragraph do |p|
  puts p
end