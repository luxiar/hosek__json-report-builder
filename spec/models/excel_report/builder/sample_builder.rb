require 'virtus'

module JsonReportBuilder
  module ExcelReport
    module Builder
      class SampleBuilder < Base
        class SampleData
          include Virtus.model
            attribute :name,       String
            attribute :name_short, String
            attribute :size,       Float
            attribute :per_price,  Float
            attribute :zip_code,   String
            attribute :address,    String
            attribute :recital,    String
        end

        def sample
          @fixtures = File.dirname(__FILE__) + '/../../../fixtures/'
          @excel_object.template_file_name = @fixtures + 'sample.xlsx'

          create_sheet1
          create_sheet2
          create_sheet3
          create_sheet4
        end

        private

        def create_sheet1
          sheet = Object::SheetObject.new(
            template_sheet_name: 'template1',
            output_sheet_name: 'data1',
            new_name: 'test_sheet1',
            print_col_index_start: 0,
            print_col_index_end: 10,
            print_row_index_start: 0,
            print_row_index_end: 10
          )
          @excel_object.sheets << sheet

          sheet.rows << Object::RowObject.new(row_index: 0, row_index_template: 0)
          sheet.rows << Object::RowObject.new(row_index: 1, row_index_template: 1)
          sheet.rows << Object::RowObject.new(row_index: 2, row_index_template: 2)
          sheet.rows << Object::RowObject.new(row_index: 3, row_index_template: 3)
          sheet.rows << Object::RowObject.new(row_index: 4, row_index_template: 4)
          sheet.merges << Object::MergeObject.new(start_row_index: 0, start_col_index: 5, end_row_index: 0, end_col_index: 6)
          sheet.merges << Object::MergeObject.new(start_row_index: 3, start_col_index: 3, end_row_index: 3, end_col_index: 4)
          sheet.merges << Object::MergeObject.new(start_row_index: 4, start_col_index: 1, end_row_index: 4, end_col_index: 2)
          sheet.merges << Object::MergeObject.new(start_row_index: 4, start_col_index: 3, end_row_index: 4, end_col_index: 6)

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
            row = Object::RowObject.new(row_index: row_index, row_index_template: 5)
            row.cols << Object::ColObject.new(col_index: 0, value: idx + 1, type: Object::ColObject::TYPE_DOUBLE)
            row.cols << Object::ColObject.new(col_index: 1, value: record.name)
            row.cols << Object::ColObject.new(col_index: 2, value: record.name_short)
            row.cols << Object::ColObject.new(col_index: 3, value: record.size, type: Object::ColObject::TYPE_DOUBLE)
            row.cols << Object::ColObject.new(col_index: 5, value: record.per_price, type: Object::ColObject::TYPE_DOUBLE)
            sheet.rows << row
            sheet.merges << Object::MergeObject.new(start_row_index: row_index, start_col_index: 3, end_row_index: row_index, end_col_index: 4)

            row_index = 6 + idx1
            row = Object::RowObject.new(row_index: row_index, row_index_template: 6)
            row.cols << Object::ColObject.new(col_index: 0, value: record.zip_code)
            row.cols << Object::ColObject.new(col_index: 1, value: record.address)
            row.cols << Object::ColObject.new(col_index: 3, value: record.recital)
            row.row_break = true
            sheet.rows << row
            sheet.merges << Object::MergeObject.new(start_row_index: row_index, start_col_index: 1, end_row_index: row_index, end_col_index: 2)
            sheet.merges << Object::MergeObject.new(start_row_index: row_index, start_col_index: 3, end_row_index: row_index, end_col_index: 6)
          end
        end

        def create_sheet2
          sheet = Object::SheetObject.new(
            template_sheet_name: 'template1',
            output_sheet_name: 'data2'
          )
          @excel_object.sheets << sheet

          sheet.rows << Object::RowObject.new(row_index: 0, row_index_template: 0)
          sheet.rows << Object::RowObject.new(row_index: 1, row_index_template: 1)

          # paste shrink size
          row = Object::RowObject.new(row_index: 3, row_index_template: 3)
          row.cols << Object::ColObject.new(col_index: 0, value: @fixtures + 'flower.jpg', type: Object::ColObject::TYPE_IMAGE, pic_row_index_e: 5, pic_col_index_e: 2)
          # paste original size
          row.cols << Object::ColObject.new(col_index: 4, value: @fixtures + 'flower.jpg', type: Object::ColObject::TYPE_IMAGE)
          sheet.rows << row

          # png too
          row = Object::RowObject.new(row_index: 31, row_index_template: 3)
          row.cols << Object::ColObject.new(col_index: 0, value: @fixtures + 'image.png', type: Object::ColObject::TYPE_IMAGE)
          sheet.rows << row
        end

        def create_sheet3
          sheet = Object::SheetObject.new(
            template_sheet_name: 'template2',
            output_sheet_name: '',
            new_name: 'fromT'
          )
          @excel_object.sheets << sheet

          sheet.rows << Object::RowObject.new(row_index: 0, row_index_template: 0)
          sheet.rows << Object::RowObject.new(row_index: 1, row_index_template: 1)

          # paste shrink size
          row = Object::RowObject.new(row_index: 3, row_index_template: 3)
          row.cols << Object::ColObject.new(col_index: 0, value: @fixtures + 'flower.jpg', type: Object::ColObject::TYPE_IMAGE, pic_row_index_e: 5, pic_col_index_e: 2)
          # paste original size
          row.cols << Object::ColObject.new(col_index: 4, value: @fixtures + 'flower.jpg', type: Object::ColObject::TYPE_IMAGE)
          sheet.rows << row

          # png too
          row = Object::RowObject.new(row_index: 31, row_index_template: 3)
          row.cols << Object::ColObject.new(col_index: 0, value: @fixtures + 'image.png', type: Object::ColObject::TYPE_IMAGE)
          sheet.rows << row
        end

        def create_sheet4
          sheet = Object::SheetObject.new(
            template_sheet_name: 'template3',
            output_sheet_name: 'data3'
          )
          @excel_object.sheets << sheet

          sheet.rows << Object::RowObject.new(row_index: 0, row_index_template: 0)
          sheet.rows << Object::RowObject.new(row_index: 1, row_index_template: 1)
          sheet.rows << Object::RowObject.new(row_index: 2, row_index_template: 2)
          row = Object::RowObject.new(row_index: 3, row_index_template: 3)
          12.times do |t|
            row.cols << Object::ColObject.new(col_index: t + 1, value: t + 1, copy_row_index: 3, copy_col_index: 1)
          end
          row.cols << Object::ColObject.new(col_index: 13, value: '合計', copy_row_index: 3, copy_col_index: 2)
          sheet.rows << row

          row = Object::RowObject.new(row_index: 4, row_index_template: 4)
          row.cols << Object::ColObject.new(col_index: 0, value: 1)
          12.times do |t|
            row.cols << Object::ColObject.new(col_index: t + 1, value: Random.rand(1..100), copy_row_index: 4, copy_col_index: 1, type: Object::ColObject::TYPE_DOUBLE)
          end
          row.cols << Object::ColObject.new(col_index: 13, value: 'SUM(B5:M5)', copy_row_index: 4, copy_col_index: 2, type: Object::ColObject::TYPE_FORMULA)
          sheet.rows << row
        end
      end
    end
  end
end
