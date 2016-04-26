#!/usr/bin/env ruby
require "http"
require "json"

class ExcelAPIClient
  attr_accessor :work_book_uuid
  attr_accessor :work_sheet_uuid

  RED = "FF0000"
  INDIAN_RED = "CD5C5C"
  BLUE = "0000FF"
  GREEN = "008000"
  BLACK = "000000"

  # Setup the API client
  def initialize(ip, port="5000")
    @host = "#{ip}:#{port}"
    @simple_cell_data = self.create_simple_cell_data %w(Apple Bananna Orange Pear)
    @client = HTTP.headers(:accept => "*/*", "Content-Type" => "application/json", "User-Agent" => "curl/7.43.0")
  end

  def create_work_book
    uri = "http://#{@host}/workbook"
    body = @client.post(uri).body
    response = JSON.parse(body)
    @work_book_uuid = response["uuid"]
  end

  def create_work_sheet(title)
    json_data = { :work_book_uuid => @work_book_uuid, :title => title }
    body = @client.post("http://#{@host}/worksheet", :json => json_data).body
    response = JSON.parse(body)
    @work_sheet_uuid = response["uuid"]
  end

  def create_row(row_data=[])
    row_data = row_data.empty? ? ["Cell Info 1", "Cell Info 2", "Cell Info 3"] : row_data
    # row_data ||= ["Cell Info 1", "Cell Info 2", "Cell Info 3"]
    cell_data = self.create_cell_data row_data
    json_data = { :work_sheet_uuid => @work_sheet_uuid, :cells => cell_data }
    r = @client.post("http://#{@host}/row", :json => json_data)
    # puts r.content_type
    body = r.body
    # response = JSON.parse(body)
  end

  # Return an array of hashes
  def create_simple_cell_data(data_a)
    cell_array = []
    data_a.each do |cell_data|
      cell_array.push({ :cell_data => cell_data })
    end
    cell_array
  end

  # Return an array of hashes
  def create_cell_data(data_a)
    cell_array = []
    data_a.each do |cell_data|
      cell_array.push({ :cell_data => cell_data, :style =>  self.create_style_data })
    end
    cell_array
  end

  # Convert RGB Hex color to the color used by AXLSX
  def color(rgb_color_string)
    "FF#{rgb_color_string}"
  end

  # return style hash
  def create_style_data
    { :fg_color => color(RED), :bg_color => color(BLUE) }
  end
end


if __FILE__ == $0
  server = "192.168.62.138"
  client = ExcelAPIClient.new("10.15.20.16", "5985")
  client.create_work_book
  puts client.work_book_uuid
  client.create_work_sheet("First Worksheet")
  puts client.work_sheet_uuid
  client.create_row ["Apple", "Bananna", "Orange", "Pear", "Mandarin and Grapes"]
  client.create_row
  client.create_row ["Water Melon", "Fruit", "Orange", "Pear", "Mandarin"]
end
