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
    @simple_cell_data = self.simple_cell_data %w(Apple Bananna Orange Pear)
    @client = HTTP.headers(:accept => "*/*", "Content-Type" => "application/json", "User-Agent" => "curl/7.43.0")
    @sheet_number = 1
    @row_number = 1
  end

  def create_work_book
    uri = "http://#{@host}/workbook"
    body = @client.post(uri).body
    response = JSON.parse(body)
    @work_book_uuid = response["uuid"]
  end

  def create_work_sheet(title)
    json_data = { :work_book_uuid => @work_book_uuid, :sheet_number => @sheet_number, :title => title }
    r = @client.post("http://#{@host}/worksheet", :json => json_data)

    if r.code == 201
      # Increment sheet number on successful creation
      @sheet_number += 1
      puts "Created work sheet"
      response = JSON.parse(r.body)
      @work_sheet_uuid = response["uuid"]
    else
      puts r.body
      raise "Error response: #{r.code}"
    end
  end

  def add_row(cell_data=nil)
    cell_data ||= @simple_cell_data
    json_data = { :work_sheet_uuid => @work_sheet_uuid, :row_number => @row_number, :cells => cell_data }
    r = @client.post("http://#{@host}/row", :json => json_data)

    if r.code == 201
      # Increment row number on successful creation
      @row_number += 1
      puts "Created row"
    else
      puts r.body
      raise "Error response: #{r.code}"
    end

    # puts r.content_type
    # response = JSON.parse(r.body)
  end

  # Create default styled row data from an array of strings
  def default_row(row_data=[])
    row_data = row_data.empty? ? ["Cell Info 1", "Cell Info 2", "Cell Info 3"] : row_data
    self.styled_cell_data row_data
  end

  # Return an array of hashes
  def simple_cell_data(data_a)
    cell_array = []
    data_a.each do |cell_data|
      cell_array.push({ :cell_data => cell_data })
    end
    cell_array
  end

  # Return an array of hashes
  def styled_cell_data(data_a)
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
  client.add_row client.default_row(["Apple", "Bananna", "Orange", "Pear", "Mandarin and Grapes"])
  client.add_row
  client.add_row
  client.add_row client.default_row ["Water Melon", "Fruit", "Orange", "Pear", "Mandarin"]
end
