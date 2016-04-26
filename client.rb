#!/usr/bin/env ruby
require "http"
require "json"

class ExcelAPIClient
  attr_accessor :work_book_uuid
  attr_accessor :work_sheet_uuid

  # Setup the work_sheet_name
  def initialize(work_sheet_name = "First Worksheet")
    @name = work_sheet_name
  end

  def create_work_book
    body = HTTP.post("http://192.168.62.138:5000/workbook").body
    response = JSON.parse(body)
    @work_book_uuid = response["uuid"]
  end

  def create_work_sheet
    json_data = { :work_book_uuid => @work_book_uuid, :title => @name }
    body = HTTP.post("http://192.168.62.138:5000/worksheet", :json => json_data).body
    response = JSON.parse(body)
    @work_sheet_uuid = response["uuid"]
  end

  def create_row
    cell_data = self.create_cell_data %w(cell1data cell2data cell3data)
    json_data = { :work_sheet_uuid => @work_sheet_uuid, :cells => cell_data }
    body = HTTP.post("http://192.168.62.138:5000/row", :json => json_data).body
    response = JSON.parse(body)
  end

  # Return an array of hashes
  def create_cell_data(data_a)
    cell_array = []
    data_a.each do |cell_data|
      cell_array.push({ :cell_data => cell_data })
    end
    cell_array
  end
end


if __FILE__ == $0
  client = ExcelAPIClient.new
  client.create_work_book
  puts client.work_book_uuid
  client.create_work_sheet
  puts client.work_sheet_uuid
  puts client.create_row
end
