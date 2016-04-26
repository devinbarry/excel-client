#!/usr/bin/env ruby
require "http"
require "json"

class ExcelAPIClient
  attr_accessor :work_book_uuid
  attr_accessor :work_sheet_uuid
  attr_accessor :row_number

  RED = "FF0000"
  INDIAN_RED = "CD5C5C"
  BLUE = "0000FF"
  GREEN = "008000"
  BLACK = "000000"
  YELLOW = "FFFF00"
  ORANGE = "FFA500"

  # Setup the API client
  def initialize(ip, port="5000")
    @host = "#{ip}:#{port}"
    @simple_cell_data = self.simple_cell_data %w(Apple Bananna Orange Pear)
    @client = HTTP.headers(:accept => "*/*", "Content-Type" => "application/json", "User-Agent" => "curl/7.43.0")
    @sheet_number = 1
    @row_number = 1

    # This is used when we apply multiple styles to a row
    @style_index = 0
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
  def styled_cell_data(data_a, style=nil)
    style ||= self.red_text_blue_bg
    cell_array = []
    data_a.each do |cell_data|
      cell_array.push({ :cell_data => cell_data, :style => self.get_next_style(style) })
    end
    cell_array
  end

  # If we supplied multiple styles for one row, alternate between them
  def get_next_style(style)
    if style.kind_of?(Array)
      @style_index += 1

      # Reset index to 0 if we have used all styles
      if @style_index >= style.length
        @style_index = 0
      end

      style[@style_index]
    else
      style
    end
  end

  # Convert RGB Hex color to the color used by AXLSX
  def color(rgb_color_string)
    "FF#{rgb_color_string}"
  end

  # return style hash
  def red_text_blue_bg
    { :fg_color => color(RED), :bg_color => color(BLUE) }
  end

  # return style hash
  def green_text_bold
    { :fg_color => color(GREEN), :bg_color => color(BLACK), :font_style => 'bold' }
  end

  # return style hash
  def yellow_text_is
    { :fg_color => color(YELLOW), :font_size => 13, :font_style => 'italic strike' }
  end

  # return style hash
  def indian_red
    { :font_face => "Apple Chancery", :bg_color => color(INDIAN_RED) }
  end

  # return style hash
  def apple_orange
    { :font_face => "Apple Chancery", :fg_color => color(ORANGE), :font_size => 14, :font_style => "underline" }
  end

  def get_download_url
    "http://#{@host}/download/#{@work_book_uuid}/filename"
  end
end


if __FILE__ == $0
  fruits_1 = ["Apple", "Bananna", "Orange", "Pear", "Mandarin and Grapes"]
  fruits_2 = ["Water Melon", "Fruit", "Orange", "Pear", "Mandarin"]
  nums = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]
  names = %w(John James Pete Jan Martin Jasper Jono Paul Mark Devin)

  server = "192.168.62.138"
  client = ExcelAPIClient.new("10.15.20.16", "5985")
  client.create_work_book
  client.work_book_uuid

  # Create a worksheet with rows
  client.create_work_sheet("First Worksheet")
  client.add_row client.styled_cell_data(fruits_1, client.apple_orange)
  client.add_row
  client.add_row
  client.add_row client.default_row(fruits_2)
  client.add_row client.styled_cell_data(nums, client.yellow_text_is)

  # Create a second worksheet with further rows
  client.row_number = 1 # reset row number
  client.create_work_sheet("2nd Worksheet")
  client.add_row client.simple_cell_data(names)
  client.add_row
  client.add_row client.default_row(fruits_2)
  client.add_row
  client.add_row client.default_row(nums)
  client.add_row client.styled_cell_data(nums, client.green_text_bold)

  # Create a third worksheet where cells are not styled uniformly
  alternating = [client.green_text_bold, client.yellow_text_is, client.indian_red]
  client.row_number = 1 # reset row number
  client.create_work_sheet("Third Sheet")
  client.add_row client.simple_cell_data(names)
  client.add_row client.simple_cell_data(names)
  client.add_row client.styled_cell_data(nums, alternating)

  puts client.get_download_url
end
