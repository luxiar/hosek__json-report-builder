require 'virtus'

module JsonReportBuilder::ExcelReport::Builder
  include JsonReportBuilder::ExcelReport::Object
  class SampleBuilder < Base
    class SampleData
      include Virtus.model
      attribute :name, String
      attribute :name_short, String
      attribute :size, Float
      attribute :per_price, Float
      attribute :zip_code, String
      attribute :address, String
      attribute :recital, String
    end

    def sample
      @fixtures = File.dirname(__FILE__) + '/../../../fixtures/'
      @excel_object.template_file_name = @fixtures + 'sample.xlsx'

      create_sheet1
      create_sheet2
      create_sheet3
    end

    private

    def create_sheet1
      sheet = SheetObject.new(
        template_sheet_name: 'template1',
        output_sheet_name: 'data1',
        new_name: 'test_sheet1',
        print_col_index_start: 0,
        print_col_index_end: 10,
        print_row_index_start: 0,
        print_row_index_end: 10
      )
      @excel_object.sheets << sheet

      sheet.rows << RowObject.new(row_index: 0, row_index_template: 0)
      sheet.rows << RowObject.new(row_index: 1, row_index_template: 1)
      sheet.rows << RowObject.new(row_index: 2, row_index_template: 2)
      sheet.rows << RowObject.new(row_index: 3, row_index_template: 3)
      sheet.rows << RowObject.new(row_index: 4, row_index_template: 4)
      sheet.merges << MergeObject.new(start_row_index: 0, start_col_index: 5, end_row_index: 0, end_col_index: 6)
      sheet.merges << MergeObject.new(start_row_index: 3, start_col_index: 3, end_row_index: 3, end_col_index: 4)
      sheet.merges << MergeObject.new(start_row_index: 4, start_col_index: 1, end_row_index: 4, end_col_index: 2)
      sheet.merges << MergeObject.new(start_row_index: 4, start_col_index: 3, end_row_index: 4, end_col_index: 6)

      records = []
      records << SampleData.new(
        name: '山の上の広い場所',
        name_short: '山頂',
        size: 2231.5,
        per_price: 15.2,
        zip_code: '110-0011',
        address: '東京都町田市大山1',
        recital: ''
      )
      records << SampleData.new(
        name: '川の近くの埋立地',
        name_short: '埋立地',
        size: 122.2,
        per_price: 8,
        zip_code: '194-0022',
        address: '東京都町田市川広場34',
        recital: '水がきれい'
      )
      records.each_with_index do |record, idx|
        idx1 = idx * 2
        row_index = 5 + idx1
        row = RowObject.new(row_index: row_index, row_index_template: 5)
        row.cols << ColObject.new(col_index: 0, value: idx + 1, type: ColObject::TYPE_DOUBLE)
        row.cols << ColObject.new(col_index: 1, value: record.name)
        row.cols << ColObject.new(col_index: 2, value: record.name_short)
        row.cols << ColObject.new(col_index: 3, value: record.size, type: ColObject::TYPE_DOUBLE)
        row.cols << ColObject.new(col_index: 5, value: record.per_price, type: ColObject::TYPE_DOUBLE)
        sheet.rows << row
        sheet.merges << MergeObject.new(start_row_index: row_index, start_col_index: 3, end_row_index: row_index, end_col_index: 4)

        row_index = 6 + idx1
        row = RowObject.new(row_index: row_index, row_index_template: 6)
        row.cols << ColObject.new(col_index: 0, value: record.zip_code)
        row.cols << ColObject.new(col_index: 1, value: record.address)
        row.cols << ColObject.new(col_index: 3, value: record.recital)
        row.row_break = true
        sheet.rows << row
        sheet.merges << MergeObject.new(start_row_index: row_index, start_col_index: 1, end_row_index: row_index, end_col_index: 2)
        sheet.merges << MergeObject.new(start_row_index: row_index, start_col_index: 3, end_row_index: row_index, end_col_index: 6)
      end
    end

    def create_sheet2
      sheet = SheetObject.new(
        template_sheet_name: 'template1',
        output_sheet_name: 'data2'
      )
      @excel_object.sheets << sheet

      sheet.rows << RowObject.new(row_index: 0, row_index_template: 0)
      sheet.rows << RowObject.new(row_index: 1, row_index_template: 1)

      # paste shrink size
      row = RowObject.new(row_index: 3, row_index_template: 3)
      row.cols << ColObject.new(col_index: 0, value: @fixtures + 'flower.jpg', type: ColObject::TYPE_IMAGE, pic_row_index_e: 5, pic_col_index_e: 2)
      # paste original size
      row.cols << ColObject.new(col_index: 4, value: @fixtures + 'flower.jpg', type: ColObject::TYPE_IMAGE)
      sheet.rows << row

      # png too
      row = RowObject.new(row_index: 31, row_index_template: 3)
      row.cols << ColObject.new(col_index: 0, value: @fixtures + 'image.png', type: ColObject::TYPE_IMAGE)
      sheet.rows << row
    end

    def create_sheet3
      sheet = SheetObject.new(
        template_sheet_name: 'template2',
        output_sheet_name: '',
        new_name: 'fromT'
      )
      @excel_object.sheets << sheet

      sheet.rows << RowObject.new(row_index: 0, row_index_template: 0)
      sheet.rows << RowObject.new(row_index: 1, row_index_template: 1)

      # paste shrink size
      row = RowObject.new(row_index: 3, row_index_template: 3)
      row.cols << ColObject.new(col_index: 0, value: @fixtures + 'flower.jpg', type: ColObject::TYPE_IMAGE, pic_row_index_e: 5, pic_col_index_e: 2)
      # paste original size
      row.cols << ColObject.new(col_index: 4, value: @fixtures + 'flower.jpg', type: ColObject::TYPE_IMAGE)
      sheet.rows << row

      # png too
      row = RowObject.new(row_index: 31, row_index_template: 3)
      row.cols << ColObject.new(col_index: 0, value: @fixtures + 'image.png', type: ColObject::TYPE_IMAGE)
      sheet.rows << row
    end
  end
end
